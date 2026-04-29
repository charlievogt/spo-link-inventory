import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { getInMemoryHashIndex } from "../services/hashIndexStore.js";
import { getAllowlist } from "../services/duplicatesAllowlistStore.js";
import { buildReport } from "../services/duplicatesQuery.js";
import { getTenantHost } from "../services/config.js";

/**
 * CSV export of duplicate-detection signals. Admin-only.
 *
 * GET /api/duplicates/export?type=exact|stale|samename
 *
 * Each row describes one pair (stale, same-name) or one (group, file)
 * tuple (exact). No pagination — caller gets every row, suitable for
 * opening in Excel / Sheets.
 */

const TENANT_HOST = getTenantHost();

function toAbs(pathOrUrl: string): string {
  if (/^https?:\/\//i.test(pathOrUrl)) return pathOrUrl;
  return TENANT_HOST + (pathOrUrl.startsWith("/") ? pathOrUrl : "/" + pathOrUrl);
}

function csvEscape(value: string | number | undefined): string {
  if (value === undefined || value === null) return "";
  const s = String(value);
  if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function row(values: Array<string | number | undefined>): string {
  return values.map(csvEscape).join(",");
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

  const type = (request.query.get("type") ?? "exact").toLowerCase();
  if (type !== "exact" && type !== "stale" && type !== "samename") {
    return { status: 400, jsonBody: { ok: false, error: "type must be exact | stale | samename" } };
  }

  let index;
  try {
    index = await getInMemoryHashIndex();
  } catch (e) {
    context.error(`duplicates export: index load failed: ${(e as Error).message}`);
    return { status: 500, jsonBody: { ok: false, error: "Hash index unavailable" } };
  }
  if (!index) {
    return {
      status: 200,
      headers: {
        "Content-Type": "text/csv; charset=utf-8",
        "Content-Disposition": `attachment; filename="duplicates-${type}-empty.csv"`,
      },
      body: "",
    };
  }

  const allowlist = await getAllowlist();
  const report = buildReport(index, allowlist);

  const lines: string[] = [];
  if (type === "exact") {
    lines.push(row([
      "sha256", "groupSize", "fileRef", "fileUrl", "fileName", "size", "sitePath", "library", "currentHashObservedAt",
    ]));
    for (const g of report.exactGroups) {
      for (const f of g.files) {
        lines.push(row([
          g.sha256,
          g.files.length,
          f.fileRef,
          toAbs(f.fileRef),
          f.fileName,
          f.size,
          f.sitePath,
          f.library,
          f.currentHashObservedAt,
        ]));
      }
    }
  } else if (type === "stale") {
    lines.push(row([
      "staleFileRef", "staleFileUrl", "staleFileName", "staleSitePath", "staleLibrary", "staleCurrentHash",
      "authoritativeFileRef", "authoritativeFileUrl", "authoritativeFileName", "authoritativeSitePath",
      "authoritativeLibrary", "authoritativeCurrentHash", "divergedAt",
    ]));
    for (const p of report.stalePairs) {
      lines.push(row([
        p.staleFileRef,
        toAbs(p.staleFileRef),
        p.staleFileName,
        p.staleSitePath,
        p.staleLibrary,
        p.staleCurrentHash,
        p.authoritativeFileRef,
        toAbs(p.authoritativeFileRef),
        p.authoritativeFileName,
        p.authoritativeSitePath,
        p.authoritativeLibrary,
        p.authoritativeCurrentHash,
        p.divergedAt,
      ]));
    }
  } else {
    lines.push(row([
      "aFileRef", "aFileUrl", "aFileName", "aSitePath", "aSize", "aSha256",
      "bFileRef", "bFileUrl", "bFileName", "bSitePath", "bSize", "bSha256",
      "sameSize",
    ]));
    for (const p of report.sameNamePairs) {
      lines.push(row([
        p.aFileRef,
        toAbs(p.aFileRef),
        p.aFileName,
        p.aSitePath,
        p.aSize,
        p.aSha256,
        p.bFileRef,
        toAbs(p.bFileRef),
        p.bFileName,
        p.bSitePath,
        p.bSize,
        p.bSha256,
        p.sameSize ? "yes" : "no",
      ]));
    }
  }

  return {
    status: 200,
    headers: {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": `attachment; filename="duplicates-${type}.csv"`,
    },
    body: lines.join("\r\n") + "\r\n",
  };
}

app.http("duplicatesExport", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "duplicates/export",
  handler: exportHandler,
});
