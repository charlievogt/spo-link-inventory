import type { FileEntry, InMemoryHashIndex, FileRefSummary } from "./hashIndexStore.js";
import { hashGroupKey } from "./hashIndexStore.js";
import { simhashDistance, NEAR_DUPLICATE_THRESHOLD } from "./textHash.js";

/**
 * Pure query logic for duplicate detection signals. Exported separately
 * from the store so endpoints and the CSV export share one
 * implementation and unit tests can cover each rule exhaustively.
 *
 * Three signals:
 *   - 🔴 Exact duplicate: 2+ files with the same `currentHash`
 *   - 🔴 Stale copy: my hash matches another file's PREVIOUS hash AND
 *     their current hash is different — I am a copy of what they used to have
 *   - ⚠ Same-name different content: pairs with same filename but
 *     different hash. Elevated to "orange" if ALSO same size.
 *
 * Allowlist suppression:
 *   - hashAllowlist: file dropped from ALL signals
 *   - pathAllowlist: any fileRef matching the glob dropped from ALL signals
 *   - nameAllowlist: filename matching the glob dropped ONLY from the
 *     same-name warning. Exact/stale are real content matches so name
 *     patterns shouldn't mask them.
 */

// --- Allowlist types ---

export interface AllowlistEntry {
  pattern?: string;
  sha256?: string;
  note: string;
  addedBy: string;
  addedAt: string;
}

export interface DuplicatesAllowlist {
  hashAllowlist: Array<{ sha256: string; note: string; addedBy: string; addedAt: string }>;
  pathAllowlist: Array<{ pattern: string; note: string; addedBy: string; addedAt: string }>;
  nameAllowlist: Array<{ pattern: string; note: string; addedBy: string; addedAt: string }>;
}

export function emptyAllowlist(): DuplicatesAllowlist {
  return { hashAllowlist: [], pathAllowlist: [], nameAllowlist: [] };
}

// --- Glob matcher (minimatch-lite) ---

/**
 * Convert a glob pattern to a RegExp. Supports:
 *   `**`  → any characters including `/`
 *   `*`   → any characters except `/`
 *   `?`   → one character except `/`
 * Case-insensitive. Other regex metachars are escaped.
 */
export function globToRegExp(pattern: string): RegExp {
  let src = "";
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === "*" && pattern[i + 1] === "*") {
      src += ".*";
      i++;
    } else if (ch === "*") {
      src += "[^/]*";
    } else if (ch === "?") {
      src += "[^/]";
    } else if (/[.+^${}()|[\]\\]/.test(ch)) {
      src += "\\" + ch;
    } else {
      src += ch;
    }
  }
  return new RegExp("^" + src + "$", "i");
}

function matchesAny(value: string, patterns: Array<{ pattern: string }>): boolean {
  for (const p of patterns) {
    if (globToRegExp(p.pattern).test(value)) return true;
  }
  return false;
}

function hashSuppressed(hash: string, allowlist: DuplicatesAllowlist): boolean {
  return allowlist.hashAllowlist.some((h) => h.sha256.toLowerCase() === hash.toLowerCase());
}

function pathSuppressed(fileRef: string, allowlist: DuplicatesAllowlist): boolean {
  return matchesAny(fileRef, allowlist.pathAllowlist);
}

function nameSuppressed(fileName: string, allowlist: DuplicatesAllowlist): boolean {
  return matchesAny(fileName, allowlist.nameAllowlist);
}

// --- Per-file lookup ---

export interface ExactDupMatch {
  kind: "exact";
  fileRef: string;
  fileName: string;
  sitePath: string;
  library: string;
  size: number;
  sha256: string;
  currentHashObservedAt: string;
}

export interface StaleMatch {
  kind: "stale";
  /** The authoritative file — the one whose PREVIOUS hash matches me. */
  authoritativeFileRef: string;
  authoritativeFileName: string;
  authoritativeSitePath: string;
  authoritativeLibrary: string;
  authoritativeCurrentHash: string;
  /** When the authoritative file last had MY hash as its current. */
  divergedAt: string;
  /** The hash I share with the authoritative file's past. */
  matchingPreviousHash: string;
}

export interface SameNameMatch {
  kind: "samename";
  fileRef: string;
  fileName: string;
  sitePath: string;
  library: string;
  size: number;
  sha256: string;
  /** True when fileName AND size match — higher-confidence warning. */
  sameSize: boolean;
}

export interface FileSignals {
  fileRef: string;
  sha256: string;
  exact: ExactDupMatch[];
  stale: StaleMatch[];
  sameName: SameNameMatch[];
}

/**
 * Compute all three signals for one target file. Applies the allowlist:
 * hash and path allowlists drop entries from every signal; name
 * allowlist only silences the same-name bucket.
 *
 * Returns `undefined` when the target isn't in the index (unknown file).
 */
export function signalsForFile(
  index: InMemoryHashIndex,
  targetFileRef: string,
  allowlist: DuplicatesAllowlist,
): FileSignals | undefined {
  const me = index.byFileRef.get(targetFileRef);
  if (!me) return undefined;

  const myHashLower = me.currentHash.toLowerCase();
  const myAlgo = me.algo;
  const myNameLower = me.fileName.toLowerCase();

  // If the target itself is suppressed by hash/path, its signals are
  // empty — but we still return a record so callers can render
  // "allowlisted" instead of pretending we have no data.
  const suppressed = hashSuppressed(me.currentHash, allowlist) || pathSuppressed(targetFileRef, allowlist);
  if (suppressed) {
    return { fileRef: targetFileRef, sha256: me.currentHash, exact: [], stale: [], sameName: [] };
  }

  const exact: ExactDupMatch[] = [];
  const sameName: SameNameMatch[] = [];

  const sameHashGroup = index.byHash.get(hashGroupKey(me.currentHash, myAlgo)) ?? [];
  for (const other of sameHashGroup) {
    if (other.fileRef === targetFileRef) continue;
    if (hashSuppressed(other.currentHash, allowlist)) continue;
    if (pathSuppressed(other.fileRef, allowlist)) continue;
    exact.push(toExactMatch(other));
  }

  const stale: StaleMatch[] = [];
  for (const [otherRef, otherEntry] of index.byFileRef) {
    if (otherRef === targetFileRef) continue;
    // Their current == mine (same hash AND same algo) is exact, not stale.
    if (otherEntry.algo === myAlgo && otherEntry.currentHash.toLowerCase() === myHashLower) continue;
    if (hashSuppressed(otherEntry.currentHash, allowlist)) continue;
    if (pathSuppressed(otherRef, allowlist)) continue;
    // Match against the other file's previous-hash list, but only entries
    // computed with the same algo as my current hash. Cross-algo matches
    // are meaningless: they're produced by different functions over
    // different inputs.
    const match = otherEntry.previousHashes.find(
      (p) => p.algo === myAlgo && p.hash.toLowerCase() === myHashLower,
    );
    if (!match) continue;
    stale.push({
      kind: "stale",
      authoritativeFileRef: otherRef,
      authoritativeFileName: otherEntry.fileName,
      authoritativeSitePath: otherEntry.sitePath,
      authoritativeLibrary: otherEntry.library,
      authoritativeCurrentHash: otherEntry.currentHash,
      divergedAt: match.observedAt,
      matchingPreviousHash: match.hash,
    });
  }

  // Same-name only when it isn't already showing up as exact (same hash + algo).
  // Also skip the same-name bucket entirely if this filename is silenced.
  if (!nameSuppressed(me.fileName, allowlist)) {
    for (const [otherRef, otherEntry] of index.byFileRef) {
      if (otherRef === targetFileRef) continue;
      if (otherEntry.fileName.toLowerCase() !== myNameLower) continue;
      if (otherEntry.algo === myAlgo && otherEntry.currentHash.toLowerCase() === myHashLower) continue;
      if (hashSuppressed(otherEntry.currentHash, allowlist)) continue;
      if (pathSuppressed(otherRef, allowlist)) continue;
      if (nameSuppressed(otherEntry.fileName, allowlist)) continue;
      sameName.push({
        kind: "samename",
        fileRef: otherRef,
        fileName: otherEntry.fileName,
        sitePath: otherEntry.sitePath,
        library: otherEntry.library,
        size: otherEntry.size,
        sha256: otherEntry.currentHash,
        sameSize: otherEntry.size === me.size,
      });
    }
  }

  return { fileRef: targetFileRef, sha256: me.currentHash, exact, stale, sameName };
}

function toExactMatch(summary: FileRefSummary): ExactDupMatch {
  return {
    kind: "exact",
    fileRef: summary.fileRef,
    fileName: summary.fileName,
    sitePath: summary.sitePath,
    library: summary.library,
    size: summary.size,
    sha256: summary.currentHash,
    currentHashObservedAt: summary.currentHashObservedAt,
  };
}

// --- Report: tenant-wide enumerations ---

export interface ExactDupGroup {
  sha256: string;
  files: Array<{
    fileRef: string;
    fileName: string;
    sitePath: string;
    library: string;
    size: number;
    currentHashObservedAt: string;
  }>;
}

export interface StalePair {
  staleFileRef: string;
  staleFileName: string;
  staleSitePath: string;
  staleLibrary: string;
  staleCurrentHash: string;
  authoritativeFileRef: string;
  authoritativeFileName: string;
  authoritativeSitePath: string;
  authoritativeLibrary: string;
  authoritativeCurrentHash: string;
  divergedAt: string;
}

export interface SameNamePair {
  aFileRef: string;
  aFileName: string;
  aSitePath: string;
  aSize: number;
  aSha256: string;
  bFileRef: string;
  bFileName: string;
  bSitePath: string;
  bSize: number;
  bSha256: string;
  sameSize: boolean;
}

/**
 * Two files whose normalized text is similar enough that their 64-bit
 * SimHashes are within a small Hamming distance. Surfaces "polish-only"
 * rename chains, copy-then-edit forks where edits were small, and
 * format-only changes where the words stayed the same.
 *
 * Excludes pairs already in the Exact tab (same content hash + algo).
 */
export interface NearDuplicatePair {
  aFileRef: string;
  aFileName: string;
  aSitePath: string;
  aLibrary: string;
  aSize: number;
  aSimhash: string;
  bFileRef: string;
  bFileName: string;
  bSitePath: string;
  bLibrary: string;
  bSize: number;
  bSimhash: string;
  /** Hamming distance between the two SimHashes (0–64). Lower is more similar. */
  hammingDistance: number;
}

/**
 * Two files whose CURRENT contents differ but whose VERSION HISTORIES
 * share at least one common ancestor — both copies were forked from a
 * common origin and have edited away independently.
 *
 * Matches on `textHash` rather than `currentHash`, because two
 * SharePoint uploads of the same source file can pick up different
 * `[trash]/*` and `customXml/_rels/*` parts that aren't fully covered
 * by the OOXML content-aware hash. The text-content hash is robust to
 * those differences.
 *
 * Excludes pairs already in the Exact tab (currents already match) and
 * the Stale tab (which already covers the asymmetric case where exactly
 * one side has edited away).
 */
export interface DivergedPair {
  aFileRef: string;
  aFileName: string;
  aSitePath: string;
  aLibrary: string;
  aCurrentHash: string;
  bFileRef: string;
  bFileName: string;
  bSitePath: string;
  bLibrary: string;
  bCurrentHash: string;
  /** The shared ancestor's text-content hash (the join key). */
  sharedAncestorTextHash: string;
  /** Earliest observation of that ancestor across both files' histories. */
  ancestorObservedAt: string;
}

export interface DuplicatesReport {
  exactGroups: ExactDupGroup[];
  stalePairs: StalePair[];
  sameNamePairs: SameNamePair[];
  nearDuplicatePairs: NearDuplicatePair[];
  divergedPairs: DivergedPair[];
  totals: {
    exactGroups: number;
    staleFiles: number;
    sameNamePairs: number;
    nearDuplicatePairs: number;
    divergedPairs: number;
  };
}

export function buildReport(
  index: InMemoryHashIndex,
  allowlist: DuplicatesAllowlist,
): DuplicatesReport {
  const exactGroups: ExactDupGroup[] = [];
  // Iterating byHash means each group is implicitly scoped to one algo
  // (the key is `${algo}:${hash}` — different algos can't collide here).
  // Pull the underlying hash and algo back out of the first member so the
  // exported `sha256` field stays a clean hex string for the UI/CSV.
  for (const [, group] of index.byHash) {
    if (group.length < 2) continue;
    const head = group[0];
    if (hashSuppressed(head.currentHash, allowlist)) continue;
    const files = group
      .filter((g) => !pathSuppressed(g.fileRef, allowlist))
      .map((g) => ({
        fileRef: g.fileRef,
        fileName: g.fileName,
        sitePath: g.sitePath,
        library: g.library,
        size: g.size,
        currentHashObservedAt: g.currentHashObservedAt,
      }));
    if (files.length < 2) continue;
    exactGroups.push({ sha256: head.currentHash, files });
  }
  exactGroups.sort((a, b) => b.files.length - a.files.length);

  const stalePairs: StalePair[] = [];
  for (const [myRef, me] of index.byFileRef) {
    if (hashSuppressed(me.currentHash, allowlist)) continue;
    if (pathSuppressed(myRef, allowlist)) continue;
    const myHashLower = me.currentHash.toLowerCase();
    const myAlgo = me.algo;
    for (const [otherRef, other] of index.byFileRef) {
      if (otherRef === myRef) continue;
      if (other.algo === myAlgo && other.currentHash.toLowerCase() === myHashLower) continue;
      if (hashSuppressed(other.currentHash, allowlist)) continue;
      if (pathSuppressed(otherRef, allowlist)) continue;
      const match = other.previousHashes.find(
        (p) => p.algo === myAlgo && p.hash.toLowerCase() === myHashLower,
      );
      if (!match) continue;
      stalePairs.push({
        staleFileRef: myRef,
        staleFileName: me.fileName,
        staleSitePath: me.sitePath,
        staleLibrary: me.library,
        staleCurrentHash: me.currentHash,
        authoritativeFileRef: otherRef,
        authoritativeFileName: other.fileName,
        authoritativeSitePath: other.sitePath,
        authoritativeLibrary: other.library,
        authoritativeCurrentHash: other.currentHash,
        divergedAt: match.observedAt,
      });
    }
  }

  const sameNamePairs: SameNamePair[] = [];
  const byName = new Map<string, Array<[string, FileEntry]>>();
  for (const [ref, entry] of index.byFileRef) {
    if (nameSuppressed(entry.fileName, allowlist)) continue;
    if (hashSuppressed(entry.currentHash, allowlist)) continue;
    if (pathSuppressed(ref, allowlist)) continue;
    const key = entry.fileName.toLowerCase();
    let list = byName.get(key);
    if (!list) {
      list = [];
      byName.set(key, list);
    }
    list.push([ref, entry]);
  }
  for (const list of byName.values()) {
    if (list.length < 2) continue;
    for (let i = 0; i < list.length; i++) {
      for (let j = i + 1; j < list.length; j++) {
        const [aRef, a] = list[i];
        const [bRef, b] = list[j];
        // Skip if both sides are exact duplicates of each other under
        // their shared algo. Cross-algo matches fall through (the bytes
        // can't be compared, so the same-name signal still has value).
        if (a.algo === b.algo && a.currentHash.toLowerCase() === b.currentHash.toLowerCase()) continue;
        sameNamePairs.push({
          aFileRef: aRef,
          aFileName: a.fileName,
          aSitePath: a.sitePath,
          aSize: a.size,
          aSha256: a.currentHash,
          bFileRef: bRef,
          bFileName: b.fileName,
          bSitePath: b.sitePath,
          bSize: b.size,
          bSha256: b.currentHash,
          sameSize: a.size === b.size,
        });
      }
    }
  }

  // Near-duplicate pairs: pairwise SimHash Hamming distance ≤ threshold.
  // O(N²) over files that have a simhash, which for reasonable tenants
  // (under ~50k files with text) finishes in seconds. Excludes exact-
  // duplicate pairs (same content hash + algo) so the Near tab shows
  // only the "fuzzy" matches that the Exact tab can't surface.
  const nearDuplicatePairs: NearDuplicatePair[] = [];
  const fileEntries: Array<[string, FileEntry]> = [];
  for (const [ref, entry] of index.byFileRef) {
    if (!entry.simhash64) continue;
    if (hashSuppressed(entry.currentHash, allowlist)) continue;
    if (pathSuppressed(ref, allowlist)) continue;
    fileEntries.push([ref, entry]);
  }
  for (let i = 0; i < fileEntries.length; i++) {
    const [aRef, a] = fileEntries[i];
    for (let j = i + 1; j < fileEntries.length; j++) {
      const [bRef, b] = fileEntries[j];
      // Skip exact dupes — they belong on the Exact tab.
      if (
        a.algo === b.algo &&
        a.currentHash.toLowerCase() === b.currentHash.toLowerCase()
      ) continue;
      const d = simhashDistance(a.simhash64!, b.simhash64!);
      if (d > NEAR_DUPLICATE_THRESHOLD) continue;
      nearDuplicatePairs.push({
        aFileRef: aRef,
        aFileName: a.fileName,
        aSitePath: a.sitePath,
        aLibrary: a.library,
        aSize: a.size,
        aSimhash: a.simhash64!,
        bFileRef: bRef,
        bFileName: b.fileName,
        bSitePath: b.sitePath,
        bLibrary: b.library,
        bSize: b.size,
        bSimhash: b.simhash64!,
        hammingDistance: d,
      });
    }
  }
  // Sort closest-first so the most likely matches surface at the top.
  nearDuplicatePairs.sort((x, y) => x.hammingDistance - y.hammingDistance);

  // Diverged pairs: shared ancestor in version history, currents differ
  // (and neither current is in the other's history — that's the Stale
  // case). Inverted index by textHash → list of (fileRef, observedAt)
  // makes this O(N×H + sum(bucket²)) instead of O(N²×H²).
  const divergedPairs: DivergedPair[] = [];
  // textHash → list of contributions from any historical version that
  // produced that textHash. Currents are NOT included because we want
  // SHARED ANCESTRY, not "I am you" matches.
  const ancestorIndex = new Map<string, Array<{ fileRef: string; observedAt: string }>>();
  for (const [ref, entry] of index.byFileRef) {
    if (hashSuppressed(entry.currentHash, allowlist)) continue;
    if (pathSuppressed(ref, allowlist)) continue;
    for (const h of entry.previousHashes) {
      if (!h.textHash) continue;
      let list = ancestorIndex.get(h.textHash);
      if (!list) {
        list = [];
        ancestorIndex.set(h.textHash, list);
      }
      list.push({ fileRef: ref, observedAt: h.observedAt });
    }
  }
  // Dedup pairs across multiple shared ancestors so a chain of common
  // history doesn't surface as N rows for the same pair of files.
  const reportedDivergedPairs = new Set<string>();
  for (const [ancestorTextHash, contributors] of ancestorIndex) {
    if (contributors.length < 2) continue;
    for (let i = 0; i < contributors.length; i++) {
      for (let j = i + 1; j < contributors.length; j++) {
        const aRef = contributors[i].fileRef;
        const bRef = contributors[j].fileRef;
        if (aRef === bRef) continue;
        const a = index.byFileRef.get(aRef);
        const b = index.byFileRef.get(bRef);
        if (!a || !b) continue;
        // Same current → already in the Exact tab.
        if (
          a.algo === b.algo &&
          a.currentHash.toLowerCase() === b.currentHash.toLowerCase()
        ) continue;
        // One side's current matches the other's history under the same
        // algo → that's the asymmetric Stale case.
        const aLowered = a.currentHash.toLowerCase();
        const bLowered = b.currentHash.toLowerCase();
        const aCurrentIsBHistory = b.previousHashes.some(
          (p) => p.algo === a.algo && p.hash.toLowerCase() === aLowered,
        );
        const bCurrentIsAHistory = a.previousHashes.some(
          (p) => p.algo === b.algo && p.hash.toLowerCase() === bLowered,
        );
        if (aCurrentIsBHistory || bCurrentIsAHistory) continue;
        const pairKey = aRef < bRef ? `${aRef}\0${bRef}` : `${bRef}\0${aRef}`;
        if (reportedDivergedPairs.has(pairKey)) continue;
        reportedDivergedPairs.add(pairKey);
        const ancestorObservedAt =
          contributors[i].observedAt.localeCompare(contributors[j].observedAt) <= 0
            ? contributors[i].observedAt
            : contributors[j].observedAt;
        divergedPairs.push({
          aFileRef: aRef,
          aFileName: a.fileName,
          aSitePath: a.sitePath,
          aLibrary: a.library,
          aCurrentHash: a.currentHash,
          bFileRef: bRef,
          bFileName: b.fileName,
          bSitePath: b.sitePath,
          bLibrary: b.library,
          bCurrentHash: b.currentHash,
          sharedAncestorTextHash: ancestorTextHash,
          ancestorObservedAt,
        });
      }
    }
  }

  return {
    exactGroups,
    stalePairs,
    sameNamePairs,
    nearDuplicatePairs,
    divergedPairs,
    totals: {
      exactGroups: exactGroups.length,
      staleFiles: new Set(stalePairs.map((p) => p.staleFileRef)).size,
      sameNamePairs: sameNamePairs.length,
      nearDuplicatePairs: nearDuplicatePairs.length,
      divergedPairs: divergedPairs.length,
    },
  };
}

// --- ACL filtering on reports (non-admins) ---

/**
 * Drop any pair/group entry that references a site the caller can't
 * read. Admins skip this entirely. For exact groups, a group is kept
 * if the caller can read at least one of the member sites; members
 * from forbidden sites collapse to a "{sitePath}" placeholder with no
 * other metadata.
 */
export function filterReportByVisibility(
  report: DuplicatesReport,
  visibleSites: Set<string>,
): DuplicatesReport {
  const visible = (sitePath: string): boolean => visibleSites.has(sitePath.toLowerCase());

  const exactGroups = report.exactGroups
    .map((g) => ({
      ...g,
      files: g.files.filter((f) => visible(f.sitePath)),
    }))
    .filter((g) => g.files.length >= 2);

  const stalePairs = report.stalePairs.filter(
    (p) => visible(p.staleSitePath) && visible(p.authoritativeSitePath),
  );
  const sameNamePairs = report.sameNamePairs.filter(
    (p) => visible(p.aSitePath) && visible(p.bSitePath),
  );
  const nearDuplicatePairs = report.nearDuplicatePairs.filter(
    (p) => visible(p.aSitePath) && visible(p.bSitePath),
  );
  const divergedPairs = report.divergedPairs.filter(
    (p) => visible(p.aSitePath) && visible(p.bSitePath),
  );

  return {
    exactGroups,
    stalePairs,
    sameNamePairs,
    nearDuplicatePairs,
    divergedPairs,
    totals: {
      exactGroups: exactGroups.length,
      staleFiles: new Set(stalePairs.map((p) => p.staleFileRef)).size,
      sameNamePairs: sameNamePairs.length,
      nearDuplicatePairs: nearDuplicatePairs.length,
      divergedPairs: divergedPairs.length,
    },
  };
}
