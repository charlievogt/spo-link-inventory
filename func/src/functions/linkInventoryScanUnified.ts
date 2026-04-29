import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { randomUUID } from "node:crypto";
import { createJob, updateJob, getJob } from "../services/linkInventoryJobStore.js";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { enumerateSites } from "../services/spoSiteEnumerator.js";
import { enqueueScanMessage } from "../services/linkInventoryQueue.js";
import { enqueueDocsScanMessage } from "../services/linkInventoryDocsQueue.js";
import { MAX_FILE_BYTES_DEFAULT, MAX_FILE_BYTES_HARD_LIMIT } from "../services/spoFilesEnumerator.js";

/**
 * POST /api/link-inventory/scan-unified
 *
 * Triggers a paired page-scan + doc-scan. Both jobs share the same
 * site list and run independently against their respective queue
 * workers, but are cross-linked via `siblingJobId` so the UI can
 * display them as a single "unified scan" entry.
 *
 * Body (optional):
 *   {
 *     sites?: string[],            // explicit; overrides app setting / enumeration
 *     enumerateAll?: boolean,      // tenant-wide via SP REST search
 *     maxFileBytes?: number,       // doc scan only (default 100 MB, cap 500 MB)
 *     modifiedAfter?: string,      // doc scan only — ISO date
 *     verifyFiles?: boolean,       // both scans run their verifier paths
 *   }
 *
 * Response (202):
 *   {
 *     ok: true,
 *     pageJobId,
 *     docJobId,
 *     sitesTotal,
 *     enumerated  // true when SP REST search enumerated the tenant
 *   }
 *
 * Why two jobs instead of one merged worker: page scans and doc scans
 * have very different work profiles (single-phase vs two-phase, fast
 * enumerate vs slow per-file fanout). Pairing them rather than merging
 * keeps the workers, queues, and result shapes intact while presenting
 * a unified UX.
 */

const SITE_PATH_RE = /^\/sites\/[a-z0-9-]+$|^\/$/i;

interface UnifiedScanBody {
  sites?: string[];
  enumerateAll?: boolean;
  maxFileBytes?: number;
  modifiedAfter?: string;
  verifyFiles?: boolean;
}

function getDefaultSites(): string[] {
  const raw = process.env.LINK_INVENTORY_SCAN_SITES ?? "";
  return raw.split(",").map((s) => s.trim()).filter((s) => s.length > 0);
}

function validateSites(sites: string[]): { valid: string[]; invalid: string[] } {
  const valid: string[] = [];
  const invalid: string[] = [];
  for (const s of sites) {
    const path = s.startsWith("/") ? s : `/${s}`;
    if (SITE_PATH_RE.test(path)) valid.push(path);
    else invalid.push(s);
  }
  return { valid, invalid };
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

  let body: UnifiedScanBody = {};
  try {
    const text = await request.text();
    if (text) body = JSON.parse(text) as UnifiedScanBody;
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
  }

  // Resolve site list (same logic as the individual triggers).
  let requested: string[] = [];
  let enumerated = false;
  if (body.sites && body.sites.length > 0) {
    requested = body.sites;
  } else if (!body.enumerateAll) {
    requested = getDefaultSites();
  }
  if (requested.length === 0) {
    try {
      const sites = await enumerateSites();
      requested = sites.map((s) => s.serverRelativeUrl);
      enumerated = true;
    } catch (err) {
      return { status: 500, jsonBody: { ok: false, error: `Site enumeration failed: ${(err as Error).message}` } };
    }
  }

  const validation = validateSites(requested);
  const valid = validation.valid;
  if (valid.length === 0) {
    return { status: 400, jsonBody: { ok: false, error: "No valid site paths", invalid: validation.invalid } };
  }

  let maxFileBytes = body.maxFileBytes ?? MAX_FILE_BYTES_DEFAULT;
  if (typeof maxFileBytes !== "number" || maxFileBytes <= 0) maxFileBytes = MAX_FILE_BYTES_DEFAULT;
  maxFileBytes = Math.min(maxFileBytes, MAX_FILE_BYTES_HARD_LIMIT);

  let modifiedAfter: Date | undefined;
  if (body.modifiedAfter) {
    const d = new Date(body.modifiedAfter);
    if (!Number.isNaN(d.getTime())) modifiedAfter = d;
  }

  // Create both job rows with mutual sibling links. We allocate IDs up
  // front so we can cross-reference them in the createJob calls.
  const pageJobId = randomUUID();
  const docJobId = randomUUID();
  const caller = user.upn ?? user.userId;

  await createJob(pageJobId, valid, caller, "pages", undefined, undefined, body.verifyFiles, docJobId);
  await createJob(
    docJobId,
    valid,
    caller,
    "documents",
    0,
    {
      maxFileBytes,
      modifiedAfter: modifiedAfter?.toISOString(),
    },
    body.verifyFiles,
    pageJobId,
  );

  // Enqueue first messages for both pipelines in parallel.
  try {
    await Promise.all([
      enqueueScanMessage({ jobId: pageJobId, siteIndex: 0 }),
      enqueueDocsScanMessage({ jobId: docJobId, phase: "enumerate", siteIndex: 0 }),
    ]);
  } catch (err) {
    // Best-effort cleanup — mark both jobs failed if either enqueue fails.
    for (const jid of [pageJobId, docJobId]) {
      try {
        const j = await getJob(jid);
        if (j) {
          j.status = "failed";
          j.recentErrors = [`unified scan enqueue failed: ${(err as Error).message}`, ...j.recentErrors].slice(0, 10);
          await updateJob(j);
        }
      } catch { /* swallow */ }
    }
    return { status: 500, jsonBody: { ok: false, error: `Enqueue failed: ${(err as Error).message}` } };
  }

  context.log(
    `[scan-unified] page=${pageJobId} doc=${docJobId} across ${valid.length} sites${enumerated ? ' (enumerated)' : ''}`,
  );

  return {
    status: 202,
    jsonBody: {
      ok: true,
      pageJobId,
      docJobId,
      sitesTotal: valid.length,
      enumerated,
      invalid: validation.invalid.length > 0 ? validation.invalid : undefined,
    },
  };
}

app.http("linkInventoryScanUnified", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "link-inventory/scan-unified",
  handler,
});
