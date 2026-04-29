import { getSpoToken, SPO_ORIGIN } from "./spoTokenProvider.js";
import type { LinkClass, MalformedReason } from "./linkInventoryScanner.js";

/**
 * Structural shape the verifier works against. Both the page-scan
 * `SiteInventory` and the doc-scan aggregate sites satisfy this —
 * we don't import the concrete types so ClassifiedLink and
 * ClassifiedDocLink can both be mutated through the same code path.
 */
interface VerifiableLink {
  url: string;
  linkClass: LinkClass;
  malformedReason?: MalformedReason;
  /**
   * Optional — page-scan `ClassifiedLink` carries this. When the
   * verifier upgrades a link's reason to `file-not-found` or `both`,
   * the prior "strip the fragment" suggestion is no longer safe
   * (the id= path is 404, so the fix isn't mechanical). We clear it
   * so the canonicalization report doesn't offer a misleading diff.
   */
  suggestion?: string;
}
interface VerifiablePage {
  links: VerifiableLink[];
}
interface VerifiableSite {
  pages: VerifiablePage[];
  byClass: Record<string, number>;
}

/**
 * Opt-in post-scan verification. Walks every AllItems.aspx?id=<path>
 * link in the aggregate and HEAD-checks the referenced file against
 * SP REST. Files that return 404 cause their links to be upgraded to
 * `malformed-spo-link` (if they weren't already) with reason
 * `file-not-found` — or `both` if the heuristic had already flagged
 * a `#search=` fragment on the same URL.
 *
 * Why this is opt-in: AllItems.aspx links are a small fraction of most
 * scans, but at tenant scale you can still rack up thousands of REST
 * calls. The page-scan dialog exposes a checkbox that sets
 * `verifyFiles: true` on the job; finalizeJob calls this helper only
 * when the flag is on.
 *
 * Mutation strategy: we mutate `ClassifiedLink` objects inside the
 * aggregate in place. This lets finalizeJob run `writeResults(aggregate)`
 * unchanged — the updated classifications are already reflected.
 * Per-site `byClass` counters are recomputed after mutation.
 *
 * Never throws. Individual REST failures are swallowed (logged via
 * `onWarn`) so a throttling hiccup can't fail the whole scan.
 */

const ALL_ITEMS_RE = /\/forms\/allitems\.aspx\?/i;
/** Max concurrent HEAD requests. SP throttles aggressively on bulk metadata. */
const CONCURRENCY = 4;

export interface VerifyStats {
  candidates: number;
  checked: number;
  notFound: number;
  errors: number;
  /** ms spent in the verify sweep. */
  durationMs: number;
}

export interface VerifyOptions {
  onWarn?: (msg: string) => void;
  onProgress?: (done: number, total: number) => void;
}

interface Candidate {
  site: string;        // /sites/foo
  idPath: string;      // /sites/foo/Shared Documents/bar.pdf (decoded)
  links: VerifiableLink[];
}

function parseIdFromAllItems(url: string): { site: string; idPath: string } | undefined {
  if (!ALL_ITEMS_RE.test(url)) return undefined;
  // Strip scheme+host if any, then fragment (we don't need it for the file check)
  const noHost = url.replace(/^https?:\/\/[^/]+/i, "");
  const [beforeHash] = noHost.split("#", 1);
  const qIdx = beforeHash.indexOf("?");
  if (qIdx < 0) return undefined;
  // URLSearchParams handles &amp; if we normalize first
  const query = beforeHash.slice(qIdx + 1).replace(/&amp;/g, "&");
  let id: string | null;
  try {
    id = new URLSearchParams(query).get("id");
  } catch {
    return undefined;
  }
  if (!id) return undefined;
  let decoded: string;
  try { decoded = decodeURIComponent(id); } catch { decoded = id; }
  const siteMatch = /^(\/(?:sites|teams)\/[^/]+)\//i.exec(decoded);
  if (!siteMatch) return undefined;
  return { site: siteMatch[1], idPath: decoded };
}

/**
 * Site-scoped existence check via SP REST. Returns `true` on 200,
 * `false` on 404, `undefined` on any other error (caller treats as
 * unknown — we don't want a transient 500 to flip good links to
 * malformed).
 *
 * Uses `$select=ServerRelativeUrl` so the response is ~100 bytes
 * instead of a full file properties payload.
 */
async function headFile(site: string, idPath: string, token: string): Promise<boolean | undefined> {
  const encoded = idPath
    .split("/")
    .map((seg) => encodeURIComponent(seg))
    .join("/")
    .replace(/'/g, "''");
  const url = `${SPO_ORIGIN}${site}/_api/web/getfilebyserverrelativeurl('${encoded}')?$select=ServerRelativeUrl`;
  let res: Response;
  try {
    res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata" },
    });
  } catch {
    return undefined;
  }
  if (res.status === 200) return true;
  if (res.status === 404) return false;
  return undefined;
}

export async function verifyMalformedLinks(
  sites: VerifiableSite[],
  opts: VerifyOptions = {},
): Promise<VerifyStats> {
  const start = Date.now();
  const warn = opts.onWarn ?? (() => undefined);

  // Build the candidate set. A single (site, idPath) key may be
  // reached via many URL shapes across many pages — we dedupe and
  // keep references to every link object that points at it.
  const byKey = new Map<string, Candidate>();
  for (const site of sites) {
    for (const page of site.pages) {
      for (const link of page.links) {
        // Verify any `spo-internal` or heuristic-flagged `malformed-spo-link`
        // that's shaped like AllItems.aspx?id=...
        if (link.linkClass !== "spo-internal" && link.linkClass !== "malformed-spo-link") continue;
        const parsed = parseIdFromAllItems(link.url);
        if (!parsed) continue;
        const key = `${parsed.site}|${parsed.idPath}`;
        let c = byKey.get(key);
        if (!c) {
          c = { site: parsed.site, idPath: parsed.idPath, links: [] };
          byKey.set(key, c);
        }
        c.links.push(link);
      }
    }
  }

  const candidates = Array.from(byKey.values());
  const stats: VerifyStats = {
    candidates: candidates.length,
    checked: 0,
    notFound: 0,
    errors: 0,
    durationMs: 0,
  };
  if (candidates.length === 0) {
    stats.durationMs = Date.now() - start;
    return stats;
  }

  let token: string;
  try {
    token = await getSpoToken();
  } catch (e) {
    warn(`verifier: token fetch failed: ${(e as Error).message}`);
    stats.durationMs = Date.now() - start;
    return stats;
  }

  // Fixed-size worker pool — each worker pulls from a shared index.
  let cursor = 0;
  async function worker(): Promise<void> {
    while (true) {
      const idx = cursor++;
      if (idx >= candidates.length) return;
      const c = candidates[idx];
      const exists = await headFile(c.site, c.idPath, token);
      stats.checked++;
      if (exists === false) {
        stats.notFound++;
        for (const link of c.links) {
          const prior: MalformedReason | undefined = link.malformedReason;
          link.linkClass = "malformed-spo-link";
          link.malformedReason =
            prior === "search-fragment" ? "both" : "file-not-found";
          // Strip-fragment suggestion is unsafe when the id= file 404s.
          link.suggestion = "";
        }
      } else if (exists === undefined) {
        stats.errors++;
      }
      if (opts.onProgress && stats.checked % 25 === 0) {
        opts.onProgress(stats.checked, candidates.length);
      }
    }
  }

  await Promise.all(Array.from({ length: CONCURRENCY }, () => worker()));

  // Recompute per-site byClass to reflect the upgraded classifications.
  for (const site of sites) {
    const byClass: Record<string, number> = {};
    for (const page of site.pages) {
      for (const link of page.links) {
        byClass[link.linkClass] = (byClass[link.linkClass] ?? 0) + 1;
      }
    }
    site.byClass = byClass;
  }

  stats.durationMs = Date.now() - start;
  return stats;
}
