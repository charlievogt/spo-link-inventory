import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { randomUUID } from "node:crypto";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { enumerateFiles, type DocumentFileRef } from "../services/spoFilesEnumerator.js";
import {
  createBootstrapJob,
  writeBootstrapManifest,
} from "../services/duplicatesBootstrapJobStore.js";
import { enqueueBootstrapMessage } from "../services/duplicatesBootstrapQueue.js";

/**
 * POST /api/duplicates/bootstrap
 *
 * Queues a version-history bootstrap job for a single library. The HTTP
 * request does two things quickly (well under the 230s Function cap):
 *
 *   1. Enumerate files in the target site/library via SP REST. Filter to
 *      the requested library title, apply includeOther, respect maxFiles.
 *   2. Write the resulting file list as a manifest blob at
 *      link-inventory-results/<jobId>/bootstrap-manifest.json, create the
 *      job row, and enqueue the first worker message.
 *
 * The worker then processes one file per message, hashing every
 * version and overlaying onto the hash index — survives any single
 * library size because each message stays under the cap.
 *
 * Body:
 *   {
 *     sitePath: "/sites/hub",
 *     libraryTitle: "Documents",
 *     includeOther?: boolean,
 *     maxFiles?: number,             // default 500, hard cap 5000
 *     maxVersionsPerFile?: number    // default 20, hard cap 100
 *   }
 *
 * Response (202):
 *   { ok: true, jobId, filesTotal, sitePath, libraryTitle }
 */

interface BootstrapBody {
  sitePath?: string;
  libraryTitle?: string;
  includeOther?: boolean;
  maxFiles?: number;
  maxVersionsPerFile?: number;
}

async function handler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  let user;
  try {
    user = parseUserPrincipal(request.headers.get("authorization"));
    await requireAdmin(user);
  } catch (err) {
    if (err instanceof AuthError) return { status: err.status, jsonBody: { ok: false, error: err.message } };
    throw err;
  }

  let body: BootstrapBody = {};
  try {
    const text = await request.text();
    if (text) body = JSON.parse(text) as BootstrapBody;
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
  }
  if (!body.sitePath || !body.libraryTitle) {
    return { status: 400, jsonBody: { ok: false, error: "sitePath and libraryTitle are required" } };
  }

  const maxFiles = Math.max(1, Math.min(body.maxFiles ?? 500, 5000));
  const maxVersionsPerFile = Math.max(1, Math.min(body.maxVersionsPerFile ?? 20, 100));
  const includeOther = body.includeOther === true;

  let files: DocumentFileRef[];
  try {
    files = await enumerateFiles(body.sitePath, { includeOther });
  } catch (e) {
    return { status: 500, jsonBody: { ok: false, error: `Enumerate failed: ${(e as Error).message}` } };
  }
  const libFiles = files.filter((f) => f.library === body.libraryTitle).slice(0, maxFiles);
  if (libFiles.length === 0) {
    return {
      status: 200,
      jsonBody: { ok: true, filesTotal: 0, info: "No files found in the specified library" },
    };
  }

  const jobId = `bootstrap-${randomUUID()}`;
  await writeBootstrapManifest(jobId, libFiles);
  await createBootstrapJob({
    jobId,
    sitePath: body.sitePath,
    libraryTitle: body.libraryTitle,
    maxVersionsPerFile,
    includeOther,
    filesTotal: libFiles.length,
    caller: user.upn ?? user.userId,
  });

  try {
    await enqueueBootstrapMessage({ jobId, fileIndex: 0 });
  } catch (e) {
    return { status: 500, jsonBody: { ok: false, error: `Enqueue failed: ${(e as Error).message}` } };
  }

  context.log(
    `[bootstrap ${jobId}] queued — ${body.sitePath}/${body.libraryTitle}, ${libFiles.length} files, maxVersionsPerFile=${maxVersionsPerFile}`,
  );

  return {
    status: 202,
    jsonBody: {
      ok: true,
      jobId,
      status: "queued",
      sitePath: body.sitePath,
      libraryTitle: body.libraryTitle,
      filesTotal: libFiles.length,
      maxVersionsPerFile,
    },
  };
}

app.http("duplicatesBootstrapFromVersions", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "duplicates/bootstrap",
  handler,
});
