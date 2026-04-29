import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { randomUUID } from "node:crypto";
import {
  createJob,
  getJob,
  readFileManifest,
  writeFileManifest,
} from "../services/linkInventoryJobStore.js";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { enqueueDocsScanMessage } from "../services/linkInventoryDocsQueue.js";
import { type DocumentFileRef } from "../services/spoFilesEnumerator.js";

/**
 * Promote a preview-only doc scan to a real scan.
 *
 * POST /api/link-inventory/scan-docs/promote
 *
 * Body:
 *   {
 *     previewJobId: string,    // required — must be a completed preview-only job
 *     sites?: string[]         // optional — narrow to a subset of the preview's sites
 *   }
 *
 * Behavior:
 *   1. Load the preview job's file manifest
 *   2. Optionally filter to only files whose `site` is in the requested set
 *   3. Create a new doc-scan job (kind=documents, previewOnly=false)
 *   4. Persist the filtered manifest under the new jobId
 *   5. Enqueue the first scan-phase message
 *
 * The preview job is left untouched so the user can promote it again
 * with a different subset.
 *
 * Permission model: admin-only (same as the doc-scan trigger).
 */

interface PromoteRequestBody {
  previewJobId?: string;
  sites?: string[];
}

async function promoteHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  // Admin gate
  let user;
  try {
    user = parseUserPrincipal(request.headers.get("authorization"));
    await requireAdmin(user);
  } catch (err) {
    if (err instanceof AuthError) {
      return { status: err.status, jsonBody: { ok: false, error: err.message } };
    }
    throw err;
  }

  let body: PromoteRequestBody = {};
  try {
    const text = await request.text();
    if (text) body = JSON.parse(text) as PromoteRequestBody;
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
  }

  if (!body.previewJobId) {
    return { status: 400, jsonBody: { ok: false, error: "Missing previewJobId" } };
  }

  const previewJob = await getJob(body.previewJobId);
  if (!previewJob) {
    return { status: 404, jsonBody: { ok: false, error: `Preview job not found: ${body.previewJobId}` } };
  }
  if (previewJob.kind !== "documents") {
    return { status: 400, jsonBody: { ok: false, error: "Preview job is not a document scan" } };
  }
  if (!previewJob.previewOnly) {
    return { status: 400, jsonBody: { ok: false, error: "Source job was not run in preview mode" } };
  }
  if (previewJob.status !== "completed") {
    return { status: 409, jsonBody: { ok: false, error: `Preview job is ${previewJob.status}, not completed` } };
  }

  // Load the preview's file manifest
  const previewManifest = (await readFileManifest(body.previewJobId)) as DocumentFileRef[] | undefined;
  if (!previewManifest || previewManifest.length === 0) {
    return { status: 404, jsonBody: { ok: false, error: "Preview job has no manifest blob" } };
  }

  // Optionally narrow to a subset of sites
  const allowedSites = body.sites && body.sites.length > 0
    ? new Set(body.sites.map((s) => s.toLowerCase()))
    : undefined;
  const filteredManifest = allowedSites
    ? previewManifest.filter((f) => allowedSites.has(f.site.toLowerCase()))
    : previewManifest;

  if (filteredManifest.length === 0) {
    return {
      status: 400,
      jsonBody: {
        ok: false,
        error: "Filter excluded every file from the preview manifest — nothing to scan",
      },
    };
  }

  // Compute the subset of sites that actually have files in the filtered manifest
  const finalSites = Array.from(new Set(filteredManifest.map((f) => f.site)));

  // Create a real (non-preview) doc scan job and persist the filtered manifest
  const newJobId = randomUUID();
  await createJob(newJobId, finalSites, user.upn ?? user.userId, "documents", filteredManifest.length, {
    maxFileBytes: previewJob.maxFileBytes,
    modifiedAfter: previewJob.modifiedAfter,
    previewOnly: false,
  });
  await writeFileManifest(newJobId, filteredManifest);
  context.log(
    `[scan-docs promote] preview ${body.previewJobId} → job ${newJobId}: ${filteredManifest.length} files across ${finalSites.length} sites`,
  );

  try {
    await enqueueDocsScanMessage({ jobId: newJobId, phase: "scan", fileIndex: 0 });
  } catch (err) {
    context.error(`[scan-docs promote] enqueue failed: ${(err as Error).message}`);
    return {
      status: 500,
      jsonBody: { ok: false, error: `Failed to enqueue: ${(err as Error).message}` },
    };
  }

  return {
    status: 202,
    jsonBody: {
      ok: true,
      jobId: newJobId,
      kind: "documents",
      status: "queued",
      sitesTotal: finalSites.length,
      filesTotal: filteredManifest.length,
      previewJobId: body.previewJobId,
    },
  };
}

app.http("linkInventoryScanDocsPromote", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "link-inventory/scan-docs/promote",
  handler: promoteHandler,
});
