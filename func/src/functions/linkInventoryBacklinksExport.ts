import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { getInMemoryIndex } from "../services/backlinksIndexStore.js";
import {
  canonicalKeyForFile,
  indexToFlatRows,
  lookupBacklinks,
  toAbsoluteSpoUrl,
} from "../services/backlinksIndex.js";
import { enumerateFiles } from "../services/spoFilesEnumerator.js";

/**
 * Admin-only CSV export of backlink data.
 *
 * GET /api/link-inventory/backlinks/export
 *
 * Two modes driven by query string:
 *   - siteUrl+libraryTitle present → per-library mode. Enumerates files
 *     in the target library, looks each up in the persistent index,
 *     emits one row per (file × backlink). Zero-backlink files emit no
 *     rows.
 *   - neither present → tenant-wide mode. Flat dump of the inverted
 *     index — one row per (canonicalKey × backlink). Sort by
 *     canonicalKey in Excel to group by target file.
 *
 * Output is streamed back as text/csv so big exports don't hold the
 * whole dataset in memory.
 */

function csvQuote(value: string | undefined | null): string {
  const s = value == null ? "" : String(value);
  if (/[",\r\n]/.test(s)) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

function csvRow(values: Array<string | undefined | null>): string {
  return values.map(csvQuote).join(",") + "\r\n";
}

async function exportHandler(
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

  const siteUrl = request.query.get("siteUrl");
  const libraryTitle = request.query.get("libraryTitle");
  const perLibrary = !!(siteUrl && libraryTitle);

  const index = await getInMemoryIndex();
  if (!index) {
    return { status: 404, jsonBody: { ok: false, error: "No backlinks index has been built yet" } };
  }

  let csv: string;
  let filename: string;

  if (perLibrary) {
    // Convert siteUrl to a server-relative site path for enumerateFiles
    let sitePath: string;
    try {
      sitePath = new URL(siteUrl!).pathname.replace(/\/$/, "") || "/";
    } catch {
      return { status: 400, jsonBody: { ok: false, error: "Invalid siteUrl" } };
    }

    let files;
    try {
      files = await enumerateFiles(sitePath);
    } catch (e) {
      context.error(`backlinks export: enumerateFiles failed: ${(e as Error).message}`);
      return { status: 500, jsonBody: { ok: false, error: "Failed to enumerate library files" } };
    }
    const libFiles = files.filter((f) => f.library === libraryTitle);

    let body = csvRow([
      "fileRef", "fileUrl", "canonicalKey", "sourceKind", "sourceSite", "sourceSiteUrl", "sourceTitle", "sourceUrl", "scannedAt",
    ]);
    for (const file of libFiles) {
      const backlinks = lookupBacklinks(index, file.fileRef);
      if (backlinks.length === 0) continue;
      const fileUrl = toAbsoluteSpoUrl(file.fileRef);
      const canonicalKey = canonicalKeyForFile(file.fileRef);
      for (const b of backlinks) {
        body += csvRow([
          file.fileRef,
          fileUrl,
          canonicalKey,
          b.sourceKind,
          b.sourceSite,
          b.sourceSiteUrl,
          b.sourceTitle,
          b.sourceUrl,
          b.sourceUpdatedAt,
        ]);
      }
    }
    csv = body;
    const libSlug = libraryTitle!.replace(/[^a-z0-9]+/gi, "-").toLowerCase();
    const siteSlug = sitePath.replace(/[^a-z0-9]+/gi, "-").toLowerCase();
    filename = `backlinks-${siteSlug}-${libSlug}.csv`;
  } else {
    // Tenant-wide: flat dump of the index.
    let body = csvRow([
      "canonicalKey", "sourceKind", "sourceSite", "sourceSiteUrl", "sourceTitle", "sourceUrl", "scannedAt",
    ]);
    const rows = indexToFlatRows(index);
    for (const r of rows) {
      body += csvRow([
        r.canonicalKey,
        r.sourceKind,
        r.sourceSite,
        r.sourceSiteUrl,
        r.sourceTitle,
        r.sourceUrl,
        r.sourceUpdatedAt,
      ]);
    }
    csv = body;
    filename = `backlinks-tenant-${new Date().toISOString().slice(0, 10)}.csv`;
  }

  return {
    status: 200,
    headers: {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": `attachment; filename="${filename}"`,
    },
    body: csv,
  };
}

app.http("linkInventoryBacklinksExport", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "link-inventory/backlinks/export",
  handler: exportHandler,
});
