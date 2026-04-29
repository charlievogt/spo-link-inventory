import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import {
  mergeObservationsIntoIndex,
  inflateHashIndex,
  emptyHashIndex,
  hashGroupKey,
  type FileObservation,
  type PersistentHashIndex,
} from "../services/hashIndexStore.js";

function mkObs(partial: Partial<FileObservation> & { fileRef: string; sha256: string }): FileObservation {
  return {
    fileName: "file.pdf",
    size: 1024,
    etag: '"{etag},1"',
    sitePath: "/sites/hr",
    library: "Shared Documents",
    ...partial,
  };
}

describe("hashIndexStore.mergeObservationsIntoIndex", () => {
  it("creates a new entry for a never-seen fileRef", () => {
    const next = mergeObservationsIntoIndex(
      emptyHashIndex(),
      [mkObs({ fileRef: "/sites/hr/Shared Documents/a.pdf", sha256: "h1" })],
      "job-1",
      "2026-04-22T10:00:00Z",
    );
    const e = next.byFileRef["/sites/hr/Shared Documents/a.pdf"];
    assert.ok(e);
    assert.equal(e.currentHash, "h1");
    assert.equal(e.currentHashObservedAt, "2026-04-22T10:00:00Z");
    assert.equal(e.lastConfirmedByJobId, "job-1");
    assert.deepEqual(e.previousHashes, []);
  });

  it("leaves history untouched and bumps lastConfirmedAt when hash unchanged", () => {
    const after1 = mergeObservationsIntoIndex(
      emptyHashIndex(),
      [mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "h1" })],
      "job-1",
      "2026-01-01T00:00:00Z",
    );
    const after2 = mergeObservationsIntoIndex(
      after1,
      [mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "h1", size: 2048 })],
      "job-2",
      "2026-04-22T00:00:00Z",
    );
    const e = after2.byFileRef["/sites/hr/a.pdf"];
    assert.equal(e.currentHash, "h1");
    assert.equal(e.currentHashObservedAt, "2026-01-01T00:00:00Z"); // unchanged
    assert.equal(e.lastConfirmedAt, "2026-04-22T00:00:00Z");
    assert.equal(e.lastConfirmedByJobId, "job-2");
    assert.equal(e.size, 2048); // metadata updated
    assert.deepEqual(e.previousHashes, []);
  });

  it("rotates the previous hash into history when content changes", () => {
    const after1 = mergeObservationsIntoIndex(
      emptyHashIndex(),
      [mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "h1" })],
      "job-1",
      "2026-01-01T00:00:00Z",
    );
    const after2 = mergeObservationsIntoIndex(
      after1,
      [mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "h2" })],
      "job-2",
      "2026-04-22T00:00:00Z",
    );
    const e = after2.byFileRef["/sites/hr/a.pdf"];
    assert.equal(e.currentHash, "h2");
    assert.equal(e.currentHashObservedAt, "2026-04-22T00:00:00Z");
    assert.equal(e.previousHashes.length, 1);
    assert.equal(e.previousHashes[0].hash, "h1");
    assert.equal(e.previousHashes[0].observedAt, "2026-01-01T00:00:00Z");
    assert.equal(e.previousHashes[0].observedByJobId, "job-1");
  });

  it("leaves files not present in the scan untouched", () => {
    const base = mergeObservationsIntoIndex(
      emptyHashIndex(),
      [
        mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "h1" }),
        mkObs({ fileRef: "/sites/hr/b.pdf", sha256: "h2" }),
      ],
      "job-1",
      "2026-01-01T00:00:00Z",
    );
    const after = mergeObservationsIntoIndex(
      base,
      [mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "h1-new" })],
      "job-2",
      "2026-04-22T00:00:00Z",
    );
    assert.equal(after.byFileRef["/sites/hr/a.pdf"].currentHash, "h1-new");
    assert.equal(after.byFileRef["/sites/hr/b.pdf"].currentHash, "h2");
    assert.equal(after.byFileRef["/sites/hr/b.pdf"].lastConfirmedByJobId, "job-1");
  });

  it("caps history at 20 entries", () => {
    let idx: PersistentHashIndex = emptyHashIndex();
    for (let i = 0; i < 25; i++) {
      idx = mergeObservationsIntoIndex(
        idx,
        [mkObs({ fileRef: "/sites/hr/a.pdf", sha256: `h${i}` })],
        `job-${i}`,
        `2026-01-${String((i % 28) + 1).padStart(2, "0")}T00:00:00Z`,
      );
    }
    const e = idx.byFileRef["/sites/hr/a.pdf"];
    assert.equal(e.previousHashes.length, 20);
    assert.equal(e.currentHash, "h24");
    // Oldest retained should be h4 (h0..h3 pruned by the 20-cap)
    assert.equal(e.previousHashes[0].hash, "h4");
  });

  it("prunes history entries older than 2 years", () => {
    let idx: PersistentHashIndex = emptyHashIndex();
    // First observation: 3 years ago
    idx = mergeObservationsIntoIndex(
      idx,
      [mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "h-old" })],
      "job-old",
      "2023-01-01T00:00:00Z",
    );
    // Rotate: still before cutoff of (2026-04-22 - 2y = 2024-04-22)
    idx = mergeObservationsIntoIndex(
      idx,
      [mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "h-mid" })],
      "job-mid",
      "2023-06-01T00:00:00Z",
    );
    // Rotate again with observedAt that prunes older entries
    idx = mergeObservationsIntoIndex(
      idx,
      [mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "h-new" })],
      "job-new",
      "2026-04-22T00:00:00Z",
    );
    const e = idx.byFileRef["/sites/hr/a.pdf"];
    assert.equal(e.currentHash, "h-new");
    // h-old (2023-01) is >2y before 2026-04 → pruned.
    // h-mid (2023-06) is >2y before 2026-04 → also pruned.
    assert.equal(e.previousHashes.length, 0);
  });

  it("skips observations with empty fileRef or sha256", () => {
    const next = mergeObservationsIntoIndex(
      emptyHashIndex(),
      [
        mkObs({ fileRef: "", sha256: "h1" }),
        mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "" }),
        mkObs({ fileRef: "/sites/hr/b.pdf", sha256: "h1" }),
      ],
      "job-1",
      "2026-04-22T00:00:00Z",
    );
    assert.equal(Object.keys(next.byFileRef).length, 1);
    assert.ok(next.byFileRef["/sites/hr/b.pdf"]);
  });
});

describe("hashIndexStore.inflateHashIndex", () => {
  it("groups fileRefs sharing the same currentHash", () => {
    const persistent = mergeObservationsIntoIndex(
      emptyHashIndex(),
      [
        mkObs({ fileRef: "/sites/hr/a.pdf", sha256: "shared" }),
        mkObs({ fileRef: "/sites/finance/b.pdf", sha256: "shared" }),
        mkObs({ fileRef: "/sites/hr/c.pdf", sha256: "unique" }),
      ],
      "job-1",
      "2026-04-22T00:00:00Z",
    );
    const inflated = inflateHashIndex(persistent);
    assert.equal(inflated.byHash.size, 2);
    const sharedGroup = inflated.byHash.get(hashGroupKey("shared", "full-sha256"));
    assert.ok(sharedGroup);
    assert.equal(sharedGroup.length, 2);
    const uniqueGroup = inflated.byHash.get(hashGroupKey("unique", "full-sha256"));
    assert.ok(uniqueGroup);
    assert.equal(uniqueGroup.length, 1);
    assert.equal(inflated.byFileRef.size, 3);
  });

  it("keeps the same hex hash in separate buckets when algos differ", () => {
    // Two .docx files with the same hex digest under different algos
    // (one legacy full-sha256, one new ooxml-content-v1) must NOT collide.
    const persistent = mergeObservationsIntoIndex(
      emptyHashIndex(),
      [
        mkObs({ fileRef: "/sites/hr/a.docx", sha256: "abc123", algo: "full-sha256" }),
        mkObs({ fileRef: "/sites/hr/b.docx", sha256: "abc123", algo: "ooxml-content-v1" }),
      ],
      "job-1",
      "2026-04-22T00:00:00Z",
    );
    const inflated = inflateHashIndex(persistent);
    assert.equal(inflated.byHash.size, 2);
    assert.equal(inflated.byHash.get(hashGroupKey("abc123", "full-sha256"))?.length, 1);
    assert.equal(inflated.byHash.get(hashGroupKey("abc123", "ooxml-content-v1"))?.length, 1);
  });

  it("backfills missing algo on legacy persisted entries to full-sha256", () => {
    // Simulate a blob written before the algo field existed.
    const legacy: PersistentHashIndex = {
      version: 1,
      builtAt: "2026-01-01T00:00:00Z",
      byFileRef: {
        "/sites/hr/legacy.pdf": {
          // algo deliberately missing — older writers didn't set it.
          currentHash: "legacy-hash",
          fileName: "legacy.pdf",
          size: 100,
          etag: '"e,1"',
          sitePath: "/sites/hr",
          library: "Shared Documents",
          lastConfirmedAt: "2026-01-01T00:00:00Z",
          lastConfirmedByJobId: "job-pre-algo",
          currentHashObservedAt: "2026-01-01T00:00:00Z",
          previousHashes: [
            // Legacy historical hash also missing algo.
            { hash: "older", observedAt: "2025-12-01T00:00:00Z", observedByJobId: "j0" } as never,
          ],
        } as never,
      },
    };
    const inflated = inflateHashIndex(legacy);
    const entry = inflated.byFileRef.get("/sites/hr/legacy.pdf");
    assert.ok(entry);
    assert.equal(entry.algo, "full-sha256");
    assert.equal(entry.previousHashes[0].algo, "full-sha256");
    assert.ok(inflated.byHash.get(hashGroupKey("legacy-hash", "full-sha256")));
  });

  it("returns an empty in-memory index for an empty persistent one", () => {
    const inflated = inflateHashIndex(emptyHashIndex());
    assert.equal(inflated.byFileRef.size, 0);
    assert.equal(inflated.byHash.size, 0);
  });
});

describe("hashIndexStore.mergeObservationsIntoIndex — algo handling", () => {
  it("preserves the prior algo when the same algo's hash changes", () => {
    let idx = mergeObservationsIntoIndex(
      emptyHashIndex(),
      [mkObs({ fileRef: "/sites/hr/a.docx", sha256: "v1", algo: "ooxml-content-v1" })],
      "job-1",
      "2026-01-01T00:00:00Z",
    );
    idx = mergeObservationsIntoIndex(
      idx,
      [mkObs({ fileRef: "/sites/hr/a.docx", sha256: "v2", algo: "ooxml-content-v1" })],
      "job-2",
      "2026-04-01T00:00:00Z",
    );
    const e = idx.byFileRef["/sites/hr/a.docx"];
    assert.equal(e.algo, "ooxml-content-v1");
    assert.equal(e.currentHash, "v2");
    assert.equal(e.previousHashes.length, 1);
    assert.equal(e.previousHashes[0].hash, "v1");
    assert.equal(e.previousHashes[0].algo, "ooxml-content-v1");
  });

  it("replaces silently on algo upgrade without rotating the prior hash into history", () => {
    // First scan: legacy full-sha256.
    let idx = mergeObservationsIntoIndex(
      emptyHashIndex(),
      [mkObs({ fileRef: "/sites/hr/a.docx", sha256: "legacy-hash", algo: "full-sha256" })],
      "job-1",
      "2026-01-01T00:00:00Z",
    );
    // Second scan: same file, new algo. Should NOT push legacy-hash into
    // previousHashes — that would generate spurious stale matches.
    idx = mergeObservationsIntoIndex(
      idx,
      [mkObs({ fileRef: "/sites/hr/a.docx", sha256: "new-hash", algo: "ooxml-content-v1" })],
      "job-2",
      "2026-04-01T00:00:00Z",
    );
    const e = idx.byFileRef["/sites/hr/a.docx"];
    assert.equal(e.algo, "ooxml-content-v1");
    assert.equal(e.currentHash, "new-hash");
    assert.deepEqual(e.previousHashes, []);
  });
});
