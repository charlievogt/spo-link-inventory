import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import {
  buildReport,
  emptyAllowlist,
  globToRegExp,
  signalsForFile,
  filterReportByVisibility,
  type DuplicatesAllowlist,
} from "../services/duplicatesQuery.js";
import {
  emptyHashIndex,
  inflateHashIndex,
  mergeObservationsIntoIndex,
  type FileObservation,
  type PersistentHashIndex,
} from "../services/hashIndexStore.js";

function obs(partial: Partial<FileObservation> & { fileRef: string; sha256: string }): FileObservation {
  return {
    fileName: "file.pdf",
    size: 100,
    etag: '"e,1"',
    sitePath: "/sites/hr",
    library: "Shared Documents",
    ...partial,
  };
}

function mkIndex(observations: FileObservation[], scans: Array<{ jobId: string; at: string; obs: FileObservation[] }> = []): ReturnType<typeof inflateHashIndex> {
  let p: PersistentHashIndex = emptyHashIndex();
  if (observations.length > 0) {
    p = mergeObservationsIntoIndex(p, observations, "job-init", "2026-01-01T00:00:00Z");
  }
  for (const s of scans) {
    p = mergeObservationsIntoIndex(p, s.obs, s.jobId, s.at);
  }
  return inflateHashIndex(p);
}

describe("duplicatesQuery.globToRegExp", () => {
  it("matches * against a single path segment", () => {
    const re = globToRegExp("/sites/*/Templates");
    assert.ok(re.test("/sites/hr/Templates"));
    assert.ok(!re.test("/sites/hr/sub/Templates"));
  });

  it("matches ** against multiple segments", () => {
    const re = globToRegExp("/sites/**/Forms/**");
    assert.ok(re.test("/sites/hr/Docs/Forms/AllItems.aspx"));
    assert.ok(re.test("/sites/finance/Forms/template.docx"));
  });

  it("is case-insensitive", () => {
    assert.ok(globToRegExp("/Sites/HR/*").test("/sites/hr/foo.pdf"));
  });

  it("escapes regex metacharacters", () => {
    const re = globToRegExp("a+b.pdf");
    assert.ok(re.test("a+b.pdf"));
    assert.ok(!re.test("aXb.pdf"));
  });
});

describe("duplicatesQuery.signalsForFile", () => {
  it("returns undefined for an unknown fileRef", () => {
    const index = mkIndex([obs({ fileRef: "/sites/hr/a.pdf", sha256: "h1" })]);
    const s = signalsForFile(index, "/sites/hr/missing.pdf", emptyAllowlist());
    assert.equal(s, undefined);
  });

  it("flags an exact duplicate across two sites", () => {
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/a.pdf", sha256: "shared" }),
      obs({ fileRef: "/sites/finance/b.pdf", sha256: "shared", sitePath: "/sites/finance" }),
    ]);
    const s = signalsForFile(index, "/sites/hr/a.pdf", emptyAllowlist());
    assert.ok(s);
    assert.equal(s.exact.length, 1);
    assert.equal(s.exact[0].fileRef, "/sites/finance/b.pdf");
    assert.equal(s.stale.length, 0);
  });

  it("flags stale when the file's hash matches another's previous hash", () => {
    // A.pdf used to be h1, now h2. B.pdf is still at h1 → B is stale relative to A.
    const persistent = mergeObservationsIntoIndex(
      mergeObservationsIntoIndex(
        emptyHashIndex(),
        [obs({ fileRef: "/sites/hr/A.pdf", sha256: "h1" })],
        "job-1",
        "2026-01-01T00:00:00Z",
      ),
      [
        obs({ fileRef: "/sites/hr/A.pdf", sha256: "h2" }),
        obs({ fileRef: "/sites/x/B.pdf", sha256: "h1", sitePath: "/sites/x", fileName: "B.pdf" }),
      ],
      "job-2",
      "2026-04-01T00:00:00Z",
    );
    const index = inflateHashIndex(persistent);
    const s = signalsForFile(index, "/sites/x/B.pdf", emptyAllowlist());
    assert.ok(s);
    assert.equal(s.stale.length, 1);
    assert.equal(s.stale[0].authoritativeFileRef, "/sites/hr/A.pdf");
    assert.equal(s.stale[0].matchingPreviousHash, "h1");
  });

  it("flags same-name warnings and elevates when size also matches", () => {
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/Handbook.pdf", sha256: "h-hr", fileName: "Handbook.pdf", size: 500 }),
      obs({ fileRef: "/sites/it/Handbook.pdf", sha256: "h-it", sitePath: "/sites/it", fileName: "Handbook.pdf", size: 500 }),
      obs({ fileRef: "/sites/fn/Handbook.pdf", sha256: "h-fn", sitePath: "/sites/fn", fileName: "Handbook.pdf", size: 900 }),
    ]);
    const s = signalsForFile(index, "/sites/hr/Handbook.pdf", emptyAllowlist());
    assert.ok(s);
    assert.equal(s.sameName.length, 2);
    const itMatch = s.sameName.find((m) => m.fileRef === "/sites/it/Handbook.pdf");
    const fnMatch = s.sameName.find((m) => m.fileRef === "/sites/fn/Handbook.pdf");
    assert.equal(itMatch?.sameSize, true);
    assert.equal(fnMatch?.sameSize, false);
  });

  it("hash allowlist suppresses exact-duplicate counterparts", () => {
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/a.pdf", sha256: "corp-logo" }),
      obs({ fileRef: "/sites/finance/b.pdf", sha256: "corp-logo", sitePath: "/sites/finance" }),
    ]);
    const allowlist: DuplicatesAllowlist = {
      hashAllowlist: [{ sha256: "corp-logo", note: "template", addedBy: "admin", addedAt: "2026-01-01" }],
      pathAllowlist: [],
      nameAllowlist: [],
    };
    const s = signalsForFile(index, "/sites/hr/a.pdf", allowlist);
    assert.ok(s);
    assert.equal(s.exact.length, 0);
  });

  it("name allowlist suppresses same-name only (not exact/stale)", () => {
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/Handbook.pdf", sha256: "h-hr", fileName: "Handbook.pdf" }),
      obs({ fileRef: "/sites/it/Handbook.pdf", sha256: "h-it", sitePath: "/sites/it", fileName: "Handbook.pdf" }),
    ]);
    const allowlist: DuplicatesAllowlist = {
      hashAllowlist: [],
      pathAllowlist: [],
      nameAllowlist: [{ pattern: "Handbook.pdf", note: "multi-site template", addedBy: "a", addedAt: "2026-01-01" }],
    };
    const s = signalsForFile(index, "/sites/hr/Handbook.pdf", allowlist);
    assert.ok(s);
    assert.equal(s.sameName.length, 0);
  });

  it("path allowlist suppresses the target across all signals", () => {
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/Templates/logo.png", sha256: "shared" }),
      obs({ fileRef: "/sites/finance/logo.png", sha256: "shared", sitePath: "/sites/finance" }),
    ]);
    const allowlist: DuplicatesAllowlist = {
      hashAllowlist: [],
      pathAllowlist: [{ pattern: "/sites/*/Templates/**", note: "templates", addedBy: "a", addedAt: "2026-01-01" }],
      nameAllowlist: [],
    };
    const s = signalsForFile(index, "/sites/hr/Templates/logo.png", allowlist);
    assert.ok(s);
    assert.equal(s.exact.length, 0);
  });
});

describe("duplicatesQuery.buildReport", () => {
  it("emits one exact group per duplicated hash with 2+ files", () => {
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/a.pdf", sha256: "shared" }),
      obs({ fileRef: "/sites/finance/b.pdf", sha256: "shared", sitePath: "/sites/finance" }),
      obs({ fileRef: "/sites/hr/c.pdf", sha256: "unique" }),
    ]);
    const report = buildReport(index, emptyAllowlist());
    assert.equal(report.exactGroups.length, 1);
    assert.equal(report.exactGroups[0].files.length, 2);
    assert.equal(report.totals.exactGroups, 1);
  });

  it("emits stale pairs (directional)", () => {
    const persistent = mergeObservationsIntoIndex(
      mergeObservationsIntoIndex(
        emptyHashIndex(),
        [obs({ fileRef: "/sites/hr/A.pdf", sha256: "h1" })],
        "j1",
        "2026-01-01T00:00:00Z",
      ),
      [
        obs({ fileRef: "/sites/hr/A.pdf", sha256: "h2" }),
        obs({ fileRef: "/sites/x/B.pdf", sha256: "h1", sitePath: "/sites/x", fileName: "B.pdf" }),
      ],
      "j2",
      "2026-04-01T00:00:00Z",
    );
    const report = buildReport(inflateHashIndex(persistent), emptyAllowlist());
    assert.equal(report.stalePairs.length, 1);
    assert.equal(report.stalePairs[0].staleFileRef, "/sites/x/B.pdf");
    assert.equal(report.stalePairs[0].authoritativeFileRef, "/sites/hr/A.pdf");
  });
});

describe("duplicatesQuery.filterReportByVisibility", () => {
  it("drops exact group members from forbidden sites and keeps groups with 2+ visible", () => {
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/a.pdf", sha256: "s", sitePath: "/sites/hr" }),
      obs({ fileRef: "/sites/finance/b.pdf", sha256: "s", sitePath: "/sites/finance" }),
      obs({ fileRef: "/sites/secret/c.pdf", sha256: "s", sitePath: "/sites/secret" }),
    ]);
    const report = buildReport(index, emptyAllowlist());
    const filtered = filterReportByVisibility(report, new Set(["/sites/hr", "/sites/finance"]));
    assert.equal(filtered.exactGroups.length, 1);
    assert.equal(filtered.exactGroups[0].files.length, 2);
  });

  it("drops same-name pairs where either side is invisible", () => {
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/h.pdf", sha256: "a", fileName: "h.pdf", sitePath: "/sites/hr" }),
      obs({ fileRef: "/sites/secret/h.pdf", sha256: "b", fileName: "h.pdf", sitePath: "/sites/secret" }),
    ]);
    const report = buildReport(index, emptyAllowlist());
    const filtered = filterReportByVisibility(report, new Set(["/sites/hr"]));
    assert.equal(filtered.sameNamePairs.length, 0);
  });
});

describe("duplicatesQuery.buildReport — near-duplicate", () => {
  it("groups two files with simhashes within the threshold", () => {
    // Hamming distance between these two 64-bit values is 2 — well under
    // the default threshold of 3.
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/a.docx", sha256: "ha", simhash64: "0000000000000000" }),
      obs({ fileRef: "/sites/finance/b.docx", sha256: "hb", simhash64: "0000000000000003" }),
    ]);
    const report = buildReport(index, emptyAllowlist());
    assert.equal(report.nearDuplicatePairs.length, 1);
    const p = report.nearDuplicatePairs[0];
    assert.equal(p.hammingDistance, 2);
    assert.equal(p.aFileRef, "/sites/hr/a.docx");
    assert.equal(p.bFileRef, "/sites/finance/b.docx");
  });

  it("excludes pairs whose Hamming distance exceeds the threshold", () => {
    const index = mkIndex([
      // 0xff = 8 set bits → Hamming distance from 0 is 8, way over threshold 3.
      obs({ fileRef: "/sites/hr/a.docx", sha256: "ha", simhash64: "0000000000000000" }),
      obs({ fileRef: "/sites/finance/b.docx", sha256: "hb", simhash64: "00000000000000ff" }),
    ]);
    const report = buildReport(index, emptyAllowlist());
    assert.equal(report.nearDuplicatePairs.length, 0);
  });

  it("excludes exact-duplicate pairs from near-duplicate results", () => {
    // Same content hash + same algo + same simhash → Exact, not Near.
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/a.docx", sha256: "shared", simhash64: "abcd1234abcd1234" }),
      obs({ fileRef: "/sites/finance/b.docx", sha256: "shared", simhash64: "abcd1234abcd1234" }),
    ]);
    const report = buildReport(index, emptyAllowlist());
    assert.equal(report.exactGroups.length, 1);
    assert.equal(report.nearDuplicatePairs.length, 0);
  });

  it("ignores files without a simhash entirely", () => {
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/a.docx", sha256: "ha", simhash64: "0000000000000000" }),
      obs({ fileRef: "/sites/finance/b.docx", sha256: "hb" }), // no simhash — file had no extractable text
    ]);
    const report = buildReport(index, emptyAllowlist());
    assert.equal(report.nearDuplicatePairs.length, 0);
  });

  it("sorts results by Hamming distance ascending", () => {
    const index = mkIndex([
      obs({ fileRef: "/sites/hr/a.docx", sha256: "ha", simhash64: "0000000000000000" }),
      // Distance 2 from a:
      obs({ fileRef: "/sites/finance/b.docx", sha256: "hb", simhash64: "0000000000000003" }),
      // Distance 1 from a:
      obs({ fileRef: "/sites/marketing/c.docx", sha256: "hc", simhash64: "0000000000000001" }),
    ]);
    const report = buildReport(index, emptyAllowlist());
    // Pairs: (a,b)=2, (a,c)=1, (b,c)=1. Sorted ascending by distance.
    assert.equal(report.nearDuplicatePairs[0].hammingDistance, 1);
    assert.equal(report.nearDuplicatePairs[report.nearDuplicatePairs.length - 1].hammingDistance, 2);
  });
});

describe("duplicatesQuery.buildReport — diverged", () => {
  function obsWithText(partial: Partial<FileObservation> & { fileRef: string; sha256: string }): FileObservation {
    return obs(partial);
  }

  it("matches two files that share an ancestor in textHash", () => {
    // Two libraries, each rotated their own current away from a shared
    // ancestor whose text was "ancestor-text-hash".
    let p: PersistentHashIndex = emptyHashIndex();
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/hr/policy.docx", sha256: "h-old", textHash: "ancestor-text-hash" }),
      obsWithText({ fileRef: "/sites/finance/policy.docx", sha256: "f-old", textHash: "ancestor-text-hash" }),
    ], "job-original", "2024-01-01T00:00:00Z");
    // Both edited away — different new content hashes AND different text.
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/hr/policy.docx", sha256: "h-new", textHash: "hr-edited" }),
    ], "job-hr-edit", "2025-01-01T00:00:00Z");
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/finance/policy.docx", sha256: "f-new", textHash: "finance-edited" }),
    ], "job-finance-edit", "2025-06-01T00:00:00Z");
    const report = buildReport(inflateHashIndex(p), emptyAllowlist());
    assert.equal(report.divergedPairs.length, 1);
    const d = report.divergedPairs[0];
    assert.equal(d.sharedAncestorTextHash, "ancestor-text-hash");
    // Currents are different.
    assert.notEqual(d.aCurrentHash, d.bCurrentHash);
  });

  it("excludes the asymmetric case (Stale handles it)", () => {
    // Library A has the original; library B was a copy that A then edited.
    // A's history contains the original; A's current is different from B's.
    // B's current matches A's history → Stale, not Diverged.
    let p: PersistentHashIndex = emptyHashIndex();
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/hr/policy.docx", sha256: "shared-old", textHash: "shared-text" }),
      obsWithText({ fileRef: "/sites/finance/policy.docx", sha256: "shared-old", textHash: "shared-text" }),
    ], "job-shared", "2024-01-01T00:00:00Z");
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/hr/policy.docx", sha256: "hr-new", textHash: "hr-edited" }),
    ], "job-hr-edit", "2025-01-01T00:00:00Z");
    // finance side is unchanged — its current still equals the shared ancestor.
    const report = buildReport(inflateHashIndex(p), emptyAllowlist());
    assert.equal(report.divergedPairs.length, 0);
    assert.equal(report.stalePairs.length, 1);
  });

  it("excludes pairs whose currents already match (Exact handles it)", () => {
    let p: PersistentHashIndex = emptyHashIndex();
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/hr/policy.docx", sha256: "shared", textHash: "shared-text" }),
      obsWithText({ fileRef: "/sites/finance/policy.docx", sha256: "shared", textHash: "shared-text" }),
    ], "job-1", "2024-01-01T00:00:00Z");
    // Both edited and converged on the same new content (improbable in
    // reality but tests the exclusion logic).
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/hr/policy.docx", sha256: "merged", textHash: "merged-text" }),
      obsWithText({ fileRef: "/sites/finance/policy.docx", sha256: "merged", textHash: "merged-text" }),
    ], "job-2", "2025-01-01T00:00:00Z");
    const report = buildReport(inflateHashIndex(p), emptyAllowlist());
    assert.equal(report.divergedPairs.length, 0);
    assert.equal(report.exactGroups.length, 1);
  });

  it("ignores entries with no historical textHash", () => {
    // Both files have history but no text-content hashes (legacy entries).
    let p: PersistentHashIndex = emptyHashIndex();
    p = mergeObservationsIntoIndex(p, [
      obs({ fileRef: "/sites/hr/policy.docx", sha256: "old" }),
      obs({ fileRef: "/sites/finance/policy.docx", sha256: "old" }),
    ], "job-1", "2024-01-01T00:00:00Z");
    p = mergeObservationsIntoIndex(p, [
      obs({ fileRef: "/sites/hr/policy.docx", sha256: "hr-new" }),
      obs({ fileRef: "/sites/finance/policy.docx", sha256: "finance-new" }),
    ], "job-2", "2025-01-01T00:00:00Z");
    const report = buildReport(inflateHashIndex(p), emptyAllowlist());
    assert.equal(report.divergedPairs.length, 0);
  });

  it("dedups across multiple shared ancestors (one row per file pair)", () => {
    // A and B share TWO historical text-hash points. The query should
    // surface them as a single Diverged pair, not two rows.
    let p: PersistentHashIndex = emptyHashIndex();
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/hr/x.docx", sha256: "v1", textHash: "ancestor-1" }),
      obsWithText({ fileRef: "/sites/finance/x.docx", sha256: "v1", textHash: "ancestor-1" }),
    ], "j1", "2024-01-01T00:00:00Z");
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/hr/x.docx", sha256: "v2", textHash: "ancestor-2" }),
      obsWithText({ fileRef: "/sites/finance/x.docx", sha256: "v2", textHash: "ancestor-2" }),
    ], "j2", "2024-06-01T00:00:00Z");
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/hr/x.docx", sha256: "hr-final", textHash: "hr-edit" }),
    ], "j3", "2025-01-01T00:00:00Z");
    p = mergeObservationsIntoIndex(p, [
      obsWithText({ fileRef: "/sites/finance/x.docx", sha256: "finance-final", textHash: "finance-edit" }),
    ], "j4", "2025-06-01T00:00:00Z");
    const report = buildReport(inflateHashIndex(p), emptyAllowlist());
    assert.equal(report.divergedPairs.length, 1);
  });
});
