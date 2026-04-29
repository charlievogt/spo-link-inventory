import { normalizeUrl } from "./urlNormalizer.js";
import { getTenantHost } from "./config.js";
import type { ISiteInventoryShape } from "../functions/linkInventoryReplace.types.js";

/**
 * Inverted index over a scan aggregate: for each target URL (keyed by
 * its canonicalKey, normalized to path-only form), the list of source
 * pages or source documents that link to it. Used by the backlinks
 * endpoint, the rebuild flow, and the CSV export.
 *
 * This module produces an *in-memory* index from one scan aggregate at
 * a time. Persistent cross-job storage lives in backlinksIndexStore.ts
 * which uses these helpers plus its own merge logic.
 *
 * Key design — unified path-only key space:
 *   normalizeUrl() emits canonical keys with or without a host prefix
 *   depending on whether the source link was absolute or relative. We
 *   strip the tenant host at index-build time and at lookup time, so
 *   any addressing form of the same file resolves to the same key.
 *
 * Non-SPO canonical keys (on-prem, external) keep their host-prefixed
 * form — they won't collide with SPO file lookups and they're still
 * useful for the tenant-wide CSV export.
 */

export type SourceKind = "page" | "document";

export interface BacklinkEntry {
  /** Whether the source is a site page or a document. */
  sourceKind: SourceKind;
  /** Short site label derived from the path, e.g. "hr" from "/sites/hr". */
  sourceSite: string;
  /** Lowercased server-relative site path. For the ACL filter. */
  sourceSitePath: string;
  /** Fully-qualified source site URL for display. */
  sourceSiteUrl: string;
  /** Page title or document file name. */
  sourceTitle: string;
  /** Fully-qualified source URL for display (page URL or document URL). */
  sourceUrl: string;
  /**
   * ISO timestamp of the scan that produced this entry. Persistent
   * index includes this per-entry so mixed-age data (old page-sourced
   * + new doc-sourced for the same site) keeps per-entry staleness.
   */
  sourceUpdatedAt: string;
}

export interface BacklinksIndex {
  /** Diagnostic string — either a jobId or "(persistent)". */
  jobId: string;
  /** Top-level "freshest" timestamp — max of all sourceUpdatedAt values. */
  scannedAt: string;
  byCanonicalKey: Map<string, BacklinkEntry[]>;
}

const TENANT_HOST = getTenantHost();
const TENANT_HOST_PREFIX = TENANT_HOST + "/";

function toPathOnlyKey(key: string): string {
  if (!key) return "";
  if (key.startsWith(TENANT_HOST_PREFIX)) {
    return key.slice(TENANT_HOST.length);
  }
  return key;
}

export function deriveSiteLabel(sitePath: string): string {
  const m = sitePath.match(/\/(?:sites|teams)\/([^/]+)/i);
  if (m) return m[1];
  if (sitePath === "/" || sitePath === "") return "root";
  const parts = sitePath.split("/").filter(Boolean);
  return parts[parts.length - 1] ?? "root";
}

export function toAbsoluteSpoUrl(relOrAbs: string): string {
  if (!relOrAbs) return "";
  if (/^https?:\/\//i.test(relOrAbs)) return relOrAbs;
  const path = relOrAbs.startsWith("/") ? relOrAbs : "/" + relOrAbs;
  return `https://${TENANT_HOST}${path}`;
}

/**
 * Build an in-memory index from one scan aggregate. `sourceKind`
 * discriminates page scans vs. document scans — the caller knows which
 * one it's processing. `scannedAt` is stamped onto every entry.
 */
export function buildBacklinksIndex(
  aggregate: ISiteInventoryShape[],
  jobId: string,
  scannedAt: string,
  sourceKind: SourceKind = "page",
): BacklinksIndex {
  const byCanonicalKey = new Map<string, BacklinkEntry[]>();

  for (const site of aggregate ?? []) {
    const rawSite = site.site ?? "";
    const sourceSitePath = rawSite.toLowerCase();
    const sourceSite = deriveSiteLabel(rawSite);
    const sourceSiteUrl = toAbsoluteSpoUrl(rawSite);

    for (const page of site.pages ?? []) {
      const sourceUrl = toAbsoluteSpoUrl(page.pageUrl ?? "");
      const sourceTitle = page.pageTitle ?? (sourceKind === "document" ? "(unnamed file)" : "(untitled page)");
      const entry: BacklinkEntry = {
        sourceKind,
        sourceSite,
        sourceSitePath,
        sourceSiteUrl,
        sourceTitle,
        sourceUrl,
        sourceUpdatedAt: scannedAt,
      };

      const seenForSource = new Set<string>();

      for (const link of page.links ?? []) {
        const canonicalKey = link.canonicalKey;
        if (!canonicalKey) continue;
        const key = toPathOnlyKey(canonicalKey);
        if (!key) continue;
        if (seenForSource.has(key)) continue;
        seenForSource.add(key);

        let list = byCanonicalKey.get(key);
        if (!list) {
          list = [];
          byCanonicalKey.set(key, list);
        }
        list.push(entry);
      }
    }
  }

  return { jobId, scannedAt, byCanonicalKey };
}

/**
 * Compute the path-only canonical key for a file's server-relative URL.
 * Normalizes via the full pipeline (including host prefix) then strips
 * the host, guaranteeing consistency with canonical keys emitted at
 * scan time.
 */
export function canonicalKeyForFile(serverRelativeUrl: string): string {
  if (!serverRelativeUrl) return "";
  const abs = toAbsoluteSpoUrl(serverRelativeUrl);
  const key = normalizeUrl(abs);
  return toPathOnlyKey(key);
}

export function lookupBacklinks(
  index: BacklinksIndex,
  serverRelativeUrl: string,
): BacklinkEntry[] {
  const key = canonicalKeyForFile(serverRelativeUrl);
  if (!key) return [];
  const list = index.byCanonicalKey.get(key);
  if (!list) return [];
  // Dedupe by (sourceKind, sourceUrl) — same source appearing under
  // multiple canonicalKey variants should only be listed once.
  const seen = new Set<string>();
  const out: BacklinkEntry[] = [];
  for (const e of list) {
    const k = `${e.sourceKind}|${e.sourceUrl}`;
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(e);
  }
  return out;
}

/**
 * Split backlinks into visible vs. hidden per the user's visible-site
 * set. Pass `null` for admins to short-circuit filtering.
 */
export function applyAclGate(
  entries: BacklinkEntry[],
  visibleSitePaths: Set<string> | null,
): { visible: BacklinkEntry[]; hiddenCount: number } {
  if (visibleSitePaths === null) {
    return { visible: entries, hiddenCount: 0 };
  }
  const visible: BacklinkEntry[] = [];
  let hiddenCount = 0;
  for (const e of entries) {
    if (visibleSitePaths.has(e.sourceSitePath)) {
      visible.push(e);
    } else {
      hiddenCount++;
    }
  }
  return { visible, hiddenCount };
}

/**
 * Flatten the index to CSV-ready rows for the tenant-wide export. One
 * row per (canonicalKey × backlink entry), sorted by canonicalKey.
 */
export function indexToFlatRows(
  index: BacklinksIndex,
): Array<BacklinkEntry & { canonicalKey: string }> {
  const out: Array<BacklinkEntry & { canonicalKey: string }> = [];
  const keys = [...index.byCanonicalKey.keys()].sort();
  for (const canonicalKey of keys) {
    const entries = index.byCanonicalKey.get(canonicalKey) ?? [];
    for (const e of entries) {
      out.push({ ...e, canonicalKey });
    }
  }
  return out;
}
