import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { randomUUID } from "node:crypto";
import { createJob, writeFileManifest } from "../services/linkInventoryJobStore.js";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { enumerateSites } from "../services/spoSiteEnumerator.js";
import { fetchFileMetadata, MAX_FILE_BYTES_DEFAULT, MAX_FILE_BYTES_HARD_LIMIT, type DocumentFileRef } from "../services/spoFilesEnumerator.js";
import { enqueueDocsScanMessage } from "../services/linkInventoryDocsQueue.js";

/**
 * Document scan trigger endpoint.
 *
 * POST /api/link-inventory/scan-docs
 *
 * Body (all optional):
 *   {
 *     sites?: string[],          // explicit list, overrides enumeration
 *     enumerateAll?: boolean,    // force enumeration
 *     maxFileBytes?: number,     // file-size cap (default 100 MB,
 *                                //   hard ceiling 500 MB)
 *     modifiedAfter?: string,    // ISO date — incremental mode
 *     fileRefs?: string[]        // targeted file mode (skips enumeration)
 *   }
 *
 * Two-phase pipeline:
 *   1. Resolve site list (explicit / setting / tenant enumeration via
 *      SP search — fast, runs in the HTTP handler).
 *   2. Persist scan params on the job row, enqueue an `enumerate`
 *      message for site #0. Worker walks one site per message,
 *      writing manifest fragments. When the last site is enumerated,
 *      the worker consolidates fragments and switches to scan phase.
 *   3. Scan phase opens one file per message and extracts links.
 *
 * Why split: a tenant-wide enumeration can take 10+ minutes by itself,
 * which exceeds the consumption-plan timeout. Splitting per-site keeps
 * each message under the cap and gives the UI live progress updates.
 *
 * `fileRefs` mode is the exception — it skips enumeration entirely
 * and goes straight to scan phase, since metadata fetches for a small
 * named list are fast.
 *
 * Permission model: admin-only (same as page scans).
 */

const SITE_PATH_RE = /^\/sites\/[a-z0-9-]+$|^\/$/i;

interface ScanDocsRequestBody {
  sites?: string[];
  enumerateAll?: boolean;
  maxFileBytes?: number;
  modifiedAfter?: string;
  /**
   * Targeted file mode: a list of server-relative file paths to scan.
   * When set, library enumeration is skipped entirely — we fetch
   * metadata for just these files and build the manifest directly.
   * Files that don't exist or aren't a parseable type are silently
   * skipped from the manifest.
   */
  fileRefs?: string[];
  /**
   * Preview-only mode. When true, the worker walks the enumerate
   * phase and writes a per-site/per-library file count summary, then
   * stops without scanning any file. The user can then promote the
   * preview to a real scan via /scan-docs/promote with an optional
   * narrowed site list.
   *
   * Only meaningful for site-mode scans — `fileRefs` mode has no
   * enumeration phase to preview.
   */
  previewOnly?: boolean;
  /** Opt-in SPO file verification during finalize. Same semantics as
   *  `verifyFiles` on the page-scan endpoint. */
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

async function scanDocsHandler(
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

  let body: ScanDocsRequestBody = {};
  try {
    const text = await request.text();
    if (text) body = JSON.parse(text) as ScanDocsRequestBody;
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
  }

  // Resolve maxFileBytes — clamp to hard ceiling, fall back to default
  let maxFileBytes = body.maxFileBytes ?? MAX_FILE_BYTES_DEFAULT;
  if (typeof maxFileBytes !== "number" || maxFileBytes <= 0) {
    maxFileBytes = MAX_FILE_BYTES_DEFAULT;
  }
  maxFileBytes = Math.min(maxFileBytes, MAX_FILE_BYTES_HARD_LIMIT);

  let modifiedAfter: Date | undefined;
  if (body.modifiedAfter) {
    const d = new Date(body.modifiedAfter);
    if (!Number.isNaN(d.getTime())) modifiedAfter = d;
  }

  // Two paths:
  //   A. fileRefs mode — caller pre-selected specific files. Skip
  //      enumeration entirely; fetch metadata one at a time, then jump
  //      straight to scan phase.
  //   B. Site mode — resolve the site list (fast), persist scan params,
  //      enqueue first enumerate-phase message, return immediately.

  if (body.fileRefs && body.fileRefs.length > 0) {
    context.log(`[scan-docs] targeted file mode: ${body.fileRefs.length} files`);
    const allFiles: DocumentFileRef[] = [];
    const enumErrors: string[] = [];
    const valid: string[] = [];

    for (const ref of body.fileRefs) {
      try {
        const meta = await fetchFileMetadata(ref);
        if (!meta) {
          enumErrors.push(`${ref}: not found or unsupported type`);
          continue;
        }
        if (meta.length > maxFileBytes) {
          enumErrors.push(`${ref}: ${(meta.length / 1024 / 1024).toFixed(1)} MB exceeds ${(maxFileBytes / 1024 / 1024).toFixed(0)} MB cap`);
          continue;
        }
        allFiles.push(meta);
        if (!valid.includes(meta.site)) valid.push(meta.site);
      } catch (err) {
        enumErrors.push(`${ref}: ${(err as Error).message}`);
      }
    }

    if (allFiles.length === 0) {
      return {
        status: 200,
        jsonBody: {
          ok: true,
          message: "No parseable files found across the requested files",
          enumErrors: enumErrors.length > 0 ? enumErrors : undefined,
        },
      };
    }

    const jobId = randomUUID();
    await createJob(jobId, valid, user.upn ?? user.userId, "documents", allFiles.length, {
      maxFileBytes,
      modifiedAfter: modifiedAfter?.toISOString(),
    }, body.verifyFiles);
    await writeFileManifest(jobId, allFiles);
    context.log(`[scan-docs] job ${jobId}: ${allFiles.length} pre-selected files, maxBytes=${maxFileBytes}`);

    try {
      await enqueueDocsScanMessage({ jobId, phase: "scan", fileIndex: 0 });
    } catch (err) {
      context.error(`[scan-docs] enqueue failed: ${(err as Error).message}`);
      return {
        status: 500,
        jsonBody: { ok: false, error: `Failed to enqueue: ${(err as Error).message}` },
      };
    }

    return {
      status: 202,
      jsonBody: {
        ok: true,
        jobId,
        kind: "documents",
        status: "queued",
        sitesTotal: valid.length,
        filesTotal: allFiles.length,
        maxFileBytes,
        enumErrors: enumErrors.length > 0 ? enumErrors : undefined,
      },
    };
  }

  // Site mode: resolve site list, then enqueue enumeration phase.
  let requested: string[] = [];
  if (body.sites && body.sites.length > 0) {
    requested = body.sites;
  } else if (!body.enumerateAll) {
    requested = getDefaultSites();
  }
  let enumerated = false;
  if (requested.length === 0) {
    try {
      context.log("[scan-docs] enumerating tenant via SP REST search");
      const sites = await enumerateSites();
      requested = sites.map((s) => s.serverRelativeUrl);
      enumerated = true;
    } catch (err) {
      return {
        status: 500,
        jsonBody: { ok: false, error: `Site enumeration failed: ${(err as Error).message}` },
      };
    }
  }

  const validation = validateSites(requested);
  const valid = validation.valid;
  const invalid = validation.invalid;
  if (valid.length === 0) {
    return { status: 400, jsonBody: { ok: false, error: "No valid site paths", invalid } };
  }

  // Create the job with filesTotal=0 — the enumerate phase fills it in.
  const jobId = randomUUID();
  await createJob(jobId, valid, user.upn ?? user.userId, "documents", 0, {
    maxFileBytes,
    modifiedAfter: modifiedAfter?.toISOString(),
    previewOnly: body.previewOnly === true,
  }, body.verifyFiles);
  context.log(
    `[scan-docs] job ${jobId}: enqueueing enumerate phase for ${valid.length} sites, maxBytes=${maxFileBytes}${body.previewOnly ? ' (preview-only)' : ''}`,
  );

  try {
    await enqueueDocsScanMessage({ jobId, phase: "enumerate", siteIndex: 0 });
  } catch (err) {
    context.error(`[scan-docs] enqueue failed: ${(err as Error).message}`);
    return {
      status: 500,
      jsonBody: { ok: false, error: `Failed to enqueue: ${(err as Error).message}` },
    };
  }

  return {
    status: 202,
    jsonBody: {
      ok: true,
      jobId,
      kind: "documents",
      status: "queued",
      phase: "enumerate",
      sitesTotal: valid.length,
      maxFileBytes,
      enumerated,
      invalid: invalid.length > 0 ? invalid : undefined,
    },
  };
}

app.http("linkInventoryScanDocs", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "link-inventory/scan-docs",
  handler: scanDocsHandler,
});
