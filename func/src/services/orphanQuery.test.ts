import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import {
  canonicalKeyForFile,
  type BacklinkEntry,
  type BacklinksIndex,
} from "./backlinksIndex.js";
import { findOrphans, extractPageFolder, type OrphanFile } from "./orphanQuery.js";
import type { SitePagesAssetFile } from "./spoSitePagesAssetsEnumerator.js";

/**
 * Build a synthetic in-memory backlinks index from a list of referenced URLs.
 * The exact shape of the source pages doesn't matter for findOrphans — only
 * which canonical keys exist in `byCanonicalKey`.
 */
function indexFor(referencedUrls: string[]): BacklinksIndex {
  const byCanonicalKey = new Map<string, BacklinkEntry[]>();
  const stubEntry: BacklinkEntry = {
    sourceKind: "page",
    sourceSite: "test",
    sourceSitePath: "/sites/test",
    sourceSiteUrl: "https://contoso.sharepoint.com/sites/test",
    sourceTitle: "Test Page",
    sourceUrl: "https://contoso.sharepoint.com/sites/test/SitePages/Test.aspx",
    sourceUpdatedAt: "2026-01-01T00:00:00Z",
  };
  for (const ref of referencedUrls) {
    const key = canonicalKeyForFile(ref);
    if (!key) continue;
    let list = byCanonicalKey.get(key);
    if (!list) {
      list = [];
      byCanonicalKey.set(key, list);
    }
    list.push(stubEntry);
  }
  return { jobId: "test", scannedAt: "2026-01-01T00:00:00Z", byCanonicalKey };
}

function file(
  serverRelativeUrl: string,
  overrides: Partial<SitePagesAssetFile> = {},
): SitePagesAssetFile {
  const segs = serverRelativeUrl.split("/");
  return {
    sitePath: "/sites/foo",
    serverRelativeUrl,
    name: segs[segs.length - 1],
    size: 1024,
    timeLastModified: "2026-01-01T00:00:00Z",
    ...overrides,
  };
}

function urlsOf(orphans: OrphanFile[]): string[] {
  return orphans.map((o) => o.serverRelativeUrl);
}

describe("extractPageFolder", () => {
  it("pulls the first segment under SiteAssets/SitePages", () => {
    assert.equal(
      extractPageFolder("/sites/foo/SiteAssets/SitePages/MyPage/img.png"),
      "MyPage",
    );
  });

  it("returns the GUID for custom-template page folders", () => {
    assert.equal(
      extractPageFolder(
        "/sites/foo/SiteAssets/SitePages/2f8c1d4a-7b3e-4f2a-9c6d-1a2b3c4d5e6f/img.png",
      ),
      "2f8c1d4a-7b3e-4f2a-9c6d-1a2b3c4d5e6f",
    );
  });

  it("returns the first segment even when nested deeper", () => {
    assert.equal(
      extractPageFolder("/sites/foo/SiteAssets/SitePages/PageA/Sub/img.png"),
      "PageA",
    );
  });

  it("is case-insensitive on the SiteAssets/SitePages prefix", () => {
    assert.equal(
      extractPageFolder("/sites/foo/siteassets/sitepages/Lower/img.png"),
      "Lower",
    );
  });

  it("returns empty string when the URL doesn't match the expected shape", () => {
    assert.equal(extractPageFolder("/sites/foo/Documents/img.png"), "");
    assert.equal(extractPageFolder(""), "");
  });
});

describe("findOrphans", () => {
  it("returns nothing when every file is referenced", () => {
    const files = [
      file("/sites/foo/SiteAssets/SitePages/A/a.png"),
      file("/sites/foo/SiteAssets/SitePages/B/b.png"),
    ];
    const idx = indexFor([
      "/sites/foo/SiteAssets/SitePages/A/a.png",
      "/sites/foo/SiteAssets/SitePages/B/b.png",
    ]);
    assert.equal(findOrphans(files, idx).length, 0);
  });

  it("flags files whose canonical key isn't in the index", () => {
    const files = [
      file("/sites/foo/SiteAssets/SitePages/A/a.png"),
      file("/sites/foo/SiteAssets/SitePages/orphan/lost.png"),
    ];
    const idx = indexFor(["/sites/foo/SiteAssets/SitePages/A/a.png"]);
    const orphans = findOrphans(files, idx);
    assert.deepEqual(urlsOf(orphans), [
      "/sites/foo/SiteAssets/SitePages/orphan/lost.png",
    ]);
  });

  it("flags everything when the index is empty", () => {
    const files = [
      file("/sites/foo/SiteAssets/SitePages/A/a.png"),
      file("/sites/foo/SiteAssets/SitePages/B/b.png"),
    ];
    const idx = indexFor([]);
    assert.equal(findOrphans(files, idx).length, 2);
  });

  it("matches case-insensitively (canonical keys are lowercased)", () => {
    const files = [file("/sites/FOO/SITEASSETS/SitePages/A/A.PNG")];
    const idx = indexFor(["/sites/foo/SiteAssets/SitePages/A/a.png"]);
    assert.equal(findOrphans(files, idx).length, 0);
  });

  it("matches when the reference is URL-encoded but the file URL has spaces", () => {
    // Reference came from a page (URL-encoded), file path has the literal
    // space. The normalizer produces the same canonical key for both.
    const files = [
      file("/sites/foo/SiteAssets/SitePages/Page Name/img.png"),
    ];
    const idx = indexFor([
      "/sites/foo/SiteAssets/SitePages/Page%20Name/img.png",
    ]);
    assert.equal(findOrphans(files, idx).length, 0);
  });

  it("reliability case: renamed page — folder name no longer matches page filename", () => {
    // The page was created as "OriginalName" (so its folder is `.../OriginalName/`),
    // then renamed to "NewName.aspx". CanvasContent1 still references the
    // ORIGINAL folder URL because SP doesn't rewrite asset URLs on rename.
    // Naive name-diff would flag the OriginalName folder as orphan and recycle
    // a live asset; reference scanning correctly retains it.
    const files = [
      file("/sites/foo/SiteAssets/SitePages/OriginalName/preserved.jpg"),
    ];
    const idx = indexFor([
      "/sites/foo/SiteAssets/SitePages/OriginalName/preserved.jpg",
    ]);
    assert.equal(findOrphans(files, idx).length, 0);
  });

  it("reliability case: custom-template page — folder is a GUID", () => {
    const guidFolder =
      "/sites/foo/SiteAssets/SitePages/2f8c1d4a-7b3e-4f2a-9c6d-1a2b3c4d5e6f/custom.png";
    const files = [file(guidFolder)];
    const idx = indexFor([guidFolder]);
    assert.equal(findOrphans(files, idx).length, 0);
  });

  it("reliability case: cross-folder reference — page A references file in page B's folder", () => {
    // PageCrossRef references a file that lives inside PageA's asset folder.
    // The file is NOT orphan because something points at it, even though
    // its "home" page is something else.
    const sharedAsset =
      "/sites/foo/SiteAssets/SitePages/PageA/shared-from-elsewhere.png";
    const files = [file(sharedAsset)];
    const idx = indexFor([sharedAsset]); // referenced from PageCrossRef
    assert.equal(findOrphans(files, idx).length, 0);
  });

  it("reliability case: deleted page — orphaned folder of files surfaces", () => {
    const files = [
      file("/sites/foo/SiteAssets/SitePages/DeletedPage/leftover.png"),
      file("/sites/foo/SiteAssets/SitePages/DeletedPage/another.jpg"),
    ];
    const idx = indexFor([]); // nothing references them anymore
    const orphans = findOrphans(files, idx);
    assert.equal(orphans.length, 2);
    for (const o of orphans) {
      assert.equal(o.pageFolder, "DeletedPage");
    }
  });

  it("reliability case: stray manual upload at SitePages/<custom>/...", () => {
    const stray = "/sites/foo/SiteAssets/SitePages/strays/manual-upload.pdf";
    const files = [file(stray)];
    const idx = indexFor([]);
    const orphans = findOrphans(files, idx);
    assert.equal(orphans.length, 1);
    assert.equal(orphans[0].pageFolder, "strays");
  });

  it("end-to-end: full PS-fixture scenario set produces the expected orphan diff", () => {
    // Same five-scenario fixture validated against the live dev tenant in
    // Phase 2 — re-encoded as a unit test here. Detected orphan set must be
    // exactly { DeletedPage/deleted.png, strays/manual-upload.pdf } —
    // nothing else, no false positives on rename / GUID / cross-folder.

    const normalAsset = "/sites/foo/SiteAssets/SitePages/PageA/normal.png";
    const renamedAsset =
      "/sites/foo/SiteAssets/SitePages/OriginalName/preserved.jpg";
    const guidAsset =
      "/sites/foo/SiteAssets/SitePages/2f8c1d4a-7b3e-4f2a-9c6d-1a2b3c4d5e6f/custom.png";
    const crossRefAsset =
      "/sites/foo/SiteAssets/SitePages/PageA/crossref-source.png";
    const deletedAsset =
      "/sites/foo/SiteAssets/SitePages/DeletedPage/leftover.png";
    const stray =
      "/sites/foo/SiteAssets/SitePages/strays/manual-upload.pdf";

    const idx = indexFor([
      normalAsset,
      renamedAsset, // page (now NewName.aspx) still references OriginalName folder
      guidAsset,
      crossRefAsset, // referenced from CrossRefPage but lives in PageA's folder
      // DeletedPage's asset and the stray are deliberately NOT referenced
    ]);

    const files = [
      file(normalAsset),
      file(renamedAsset),
      file(guidAsset),
      file(crossRefAsset),
      file(deletedAsset),
      file(stray),
    ];

    const orphans = findOrphans(files, idx);
    const orphanUrls = urlsOf(orphans).sort();
    assert.deepEqual(orphanUrls, [deletedAsset, stray].sort());
  });

  it("populates pageFolder, size, modified, and canonicalKey on orphan rows", () => {
    const url = "/sites/foo/SiteAssets/SitePages/MyPage/img.png";
    const files = [
      file(url, { size: 4096, timeLastModified: "2026-04-01T10:00:00Z" }),
    ];
    const orphans = findOrphans(files, indexFor([]));
    assert.equal(orphans.length, 1);
    const o = orphans[0];
    assert.equal(o.fileName, "img.png");
    assert.equal(o.pageFolder, "MyPage");
    assert.equal(o.size, 4096);
    assert.equal(o.modified, "2026-04-01T10:00:00Z");
    assert.ok(o.canonicalKey.length > 0);
  });
});
