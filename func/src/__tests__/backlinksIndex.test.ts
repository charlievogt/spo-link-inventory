import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import {
  applyAclGate,
  buildBacklinksIndex,
  canonicalKeyForFile,
  indexToFlatRows,
  lookupBacklinks,
  type BacklinkEntry,
} from "../services/backlinksIndex.js";
import type { ISiteInventoryShape } from "../functions/linkInventoryReplace.types.js";

function mkPage(
  pageId: number,
  pageTitle: string,
  pageUrl: string,
  links: Array<{ url: string; canonicalKey?: string }>,
): ISiteInventoryShape["pages"][number] {
  return {
    pageId,
    pageTitle,
    pageUrl,
    links: links.map((l) => ({
      rawUrl: l.url,
      url: l.url,
      source: "test",
      linkClass: "spo-doc",
      normalizedKey: l.canonicalKey ?? "",
      canonicalKey: l.canonicalKey,
    })),
  };
}

function mkSite(
  site: string,
  pages: ISiteInventoryShape["pages"],
): ISiteInventoryShape {
  return {
    site,
    pageCount: pages.length,
    linkCount: pages.reduce((n, p) => n + p.links.length, 0),
    pages,
  };
}

describe("backlinksIndex", () => {
  describe("buildBacklinksIndex (page kind)", () => {
    it("indexes one page linking to one target", () => {
      const agg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(1, "Benefits", "/sites/hr/SitePages/Benefits.aspx", [
            {
              url: "https://contoso.sharepoint.com/sites/hr/Shared Documents/handbook.pdf",
              canonicalKey: "contoso.sharepoint.com/sites/hr/shared documents/handbook.pdf",
            },
          ]),
        ]),
      ];
      const idx = buildBacklinksIndex(agg, "job-1", "2026-04-22T10:00:00Z", "page");
      assert.equal(idx.byCanonicalKey.size, 1);
      const entries = idx.byCanonicalKey.get("/sites/hr/shared documents/handbook.pdf");
      assert.ok(entries, "path-only key should be indexed");
      assert.equal(entries.length, 1);
      assert.equal(entries[0].sourceKind, "page");
      assert.equal(entries[0].sourceSite, "hr");
      assert.equal(entries[0].sourceSitePath, "/sites/hr");
      assert.equal(entries[0].sourceTitle, "Benefits");
      assert.equal(entries[0].sourceUpdatedAt, "2026-04-22T10:00:00Z");
    });

    it("collapses host-prefixed and path-only keys for the same target", () => {
      const agg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(1, "Page A", "/sites/hr/SitePages/A.aspx", [
            {
              url: "https://contoso.sharepoint.com/sites/hr/Shared Documents/foo.pdf",
              canonicalKey: "contoso.sharepoint.com/sites/hr/shared documents/foo.pdf",
            },
          ]),
          mkPage(2, "Page B", "/sites/hr/SitePages/B.aspx", [
            {
              url: "/sites/hr/Shared Documents/foo.pdf",
              canonicalKey: "/sites/hr/shared documents/foo.pdf",
            },
          ]),
        ]),
      ];
      const idx = buildBacklinksIndex(agg, "job-1", "2026-04-22T10:00:00Z");
      assert.equal(idx.byCanonicalKey.size, 1);
      const entries = idx.byCanonicalKey.get("/sites/hr/shared documents/foo.pdf");
      assert.ok(entries);
      assert.equal(entries.length, 2);
      assert.deepEqual(
        entries.map((e) => e.sourceTitle).sort(),
        ["Page A", "Page B"],
      );
    });

    it("dedupes within one source when the same target appears multiple times", () => {
      const agg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(1, "Multi-link", "/sites/hr/SitePages/Multi.aspx", [
            {
              url: "https://contoso.sharepoint.com/sites/hr/docs/foo.pdf",
              canonicalKey: "/sites/hr/docs/foo.pdf",
            },
            {
              url: "/sites/hr/docs/foo.pdf",
              canonicalKey: "/sites/hr/docs/foo.pdf",
            },
          ]),
        ]),
      ];
      const idx = buildBacklinksIndex(agg, "job-1", "2026-04-22T10:00:00Z");
      const entries = idx.byCanonicalKey.get("/sites/hr/docs/foo.pdf");
      assert.ok(entries);
      assert.equal(entries.length, 1);
    });

    it("skips links without a canonicalKey", () => {
      const agg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(1, "Mixed", "/sites/hr/SitePages/Mixed.aspx", [
            { url: "https://example.com/external", canonicalKey: undefined },
            { url: "/sites/hr/foo.pdf", canonicalKey: "/sites/hr/foo.pdf" },
          ]),
        ]),
      ];
      const idx = buildBacklinksIndex(agg, "job-1", "2026-04-22T10:00:00Z");
      assert.equal(idx.byCanonicalKey.size, 1);
    });

    it("preserves non-SPO canonicalKeys in their host-prefixed form", () => {
      const agg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(1, "Old-world", "/sites/hr/SitePages/Old.aspx", [
            {
              url: "https://legacy.example.com/sites/hr/file.docx",
              canonicalKey: "legacy.example.com/sites/hr/file.docx",
            },
          ]),
        ]),
      ];
      const idx = buildBacklinksIndex(agg, "job-1", "2026-04-22T10:00:00Z");
      assert.ok(idx.byCanonicalKey.has("legacy.example.com/sites/hr/file.docx"));
    });
  });

  describe("buildBacklinksIndex (document kind)", () => {
    it("tags doc-scan sources with sourceKind='document'", () => {
      // Doc scan's finalizeJob wraps files into page-shaped records
      // with pageId=0, pageTitle=fileName, pageUrl=fileRef. When we
      // pass kind='document' the index should label them accordingly.
      const agg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(0, "procedure.docx", "/sites/hr/Shared Documents/procedure.docx", [
            {
              url: "https://contoso.sharepoint.com/sites/charlie-test-site/Shared Documents/target.pdf",
              canonicalKey: "/sites/charlie-test-site/shared documents/target.pdf",
            },
          ]),
        ]),
      ];
      const idx = buildBacklinksIndex(agg, "job-doc", "2026-04-22T10:00:00Z", "document");
      const entries = idx.byCanonicalKey.get("/sites/charlie-test-site/shared documents/target.pdf");
      assert.ok(entries);
      assert.equal(entries.length, 1);
      assert.equal(entries[0].sourceKind, "document");
      assert.equal(entries[0].sourceTitle, "procedure.docx");
    });
  });

  describe("lookupBacklinks", () => {
    const agg: ISiteInventoryShape[] = [
      mkSite("/sites/hr", [
        mkPage(1, "Benefits", "/sites/hr/SitePages/Benefits.aspx", [
          {
            url: "https://contoso.sharepoint.com/sites/charlie-test-site/Shared Documents/target.pdf",
            canonicalKey: "/sites/charlie-test-site/shared documents/target.pdf",
          },
        ]),
      ]),
      mkSite("/sites/marketing", [
        mkPage(1, "Campaign", "/sites/marketing/SitePages/Campaign.aspx", [
          {
            url: "/sites/charlie-test-site/Shared Documents/target.pdf",
            canonicalKey: "/sites/charlie-test-site/shared documents/target.pdf",
          },
        ]),
      ]),
    ];
    const idx = buildBacklinksIndex(agg, "job-1", "2026-04-22T10:00:00Z");

    it("finds backlinks by a file's server-relative URL", () => {
      const results = lookupBacklinks(idx, "/sites/charlie-test-site/Shared Documents/target.pdf");
      assert.equal(results.length, 2);
    });

    it("is case-insensitive via normalizeUrl", () => {
      const results = lookupBacklinks(idx, "/Sites/Charlie-Test-Site/Shared Documents/target.pdf");
      assert.equal(results.length, 2);
    });

    it("returns empty for unknown files", () => {
      const results = lookupBacklinks(idx, "/sites/nowhere/file.pdf");
      assert.deepEqual(results, []);
    });

    it("dedupes page + doc sources pointing at the same source URL via multiple key variants", () => {
      const dup: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(1, "P1", "/sites/hr/SitePages/P1.aspx", [
            { url: "a", canonicalKey: "/sites/x/foo.pdf" },
          ]),
        ]),
      ];
      const dupIdx = buildBacklinksIndex(dup, "j", "t");
      const extras = [...(dupIdx.byCanonicalKey.get("/sites/x/foo.pdf") ?? [])];
      dupIdx.byCanonicalKey.set("contoso.sharepoint.com/sites/x/foo.pdf", extras);
      const results = lookupBacklinks(dupIdx, "/sites/x/foo.pdf");
      assert.equal(results.length, 1);
    });

    it("keeps page and doc sources separate when both link to the same target from different URLs", () => {
      // Same target, one page source and one doc source.
      const key = "/sites/charlie-test-site/shared documents/target.pdf";
      const entries: BacklinkEntry[] = [
        {
          sourceKind: "page",
          sourceSite: "hr",
          sourceSitePath: "/sites/hr",
          sourceSiteUrl: "https://contoso.sharepoint.com/sites/hr",
          sourceTitle: "Benefits",
          sourceUrl: "https://contoso.sharepoint.com/sites/hr/SitePages/Benefits.aspx",
          sourceUpdatedAt: "2026-04-22T10:00:00Z",
        },
        {
          sourceKind: "document",
          sourceSite: "hr",
          sourceSitePath: "/sites/hr",
          sourceSiteUrl: "https://contoso.sharepoint.com/sites/hr",
          sourceTitle: "procedure.docx",
          sourceUrl: "https://contoso.sharepoint.com/sites/hr/Shared Documents/procedure.docx",
          sourceUpdatedAt: "2026-04-22T10:00:00Z",
        },
      ];
      const idx2 = {
        jobId: "synthetic",
        scannedAt: "2026-04-22T10:00:00Z",
        byCanonicalKey: new Map<string, BacklinkEntry[]>([[key, entries]]),
      };
      const out = lookupBacklinks(idx2, "/sites/charlie-test-site/Shared Documents/target.pdf");
      assert.equal(out.length, 2);
      assert.deepEqual(out.map((e) => e.sourceKind).sort(), ["document", "page"]);
    });
  });

  describe("applyAclGate", () => {
    const mk = (site: string): BacklinkEntry => ({
      sourceKind: "page",
      sourceSite: site,
      sourceSitePath: `/sites/${site}`,
      sourceSiteUrl: `https://contoso.sharepoint.com/sites/${site}`,
      sourceTitle: "P",
      sourceUrl: `https://contoso.sharepoint.com/sites/${site}/SitePages/P.aspx`,
      sourceUpdatedAt: "2026-04-22T10:00:00Z",
    });
    const entries = [mk("hr"), mk("secret")];

    it("returns all entries when visibleSitePaths is null (admin)", () => {
      const { visible, hiddenCount } = applyAclGate(entries, null);
      assert.equal(visible.length, 2);
      assert.equal(hiddenCount, 0);
    });

    it("filters to only visible sites for non-admins", () => {
      const { visible, hiddenCount } = applyAclGate(entries, new Set(["/sites/hr"]));
      assert.equal(visible.length, 1);
      assert.equal(visible[0].sourceSite, "hr");
      assert.equal(hiddenCount, 1);
    });

    it("returns empty visible + full hidden count when user sees none", () => {
      const { visible, hiddenCount } = applyAclGate(entries, new Set());
      assert.equal(visible.length, 0);
      assert.equal(hiddenCount, 2);
    });
  });

  describe("canonicalKeyForFile", () => {
    it("normalizes mixed-case paths to lowercase", () => {
      const key = canonicalKeyForFile("/Sites/HR/Shared Documents/Handbook.pdf");
      assert.equal(key, "/sites/hr/shared documents/handbook.pdf");
    });

    it("decodes URL-encoded paths", () => {
      const key = canonicalKeyForFile("/sites/hr/Shared%20Documents/handbook.pdf");
      assert.equal(key, "/sites/hr/shared documents/handbook.pdf");
    });

    it("returns empty for empty input", () => {
      assert.equal(canonicalKeyForFile(""), "");
    });
  });

  describe("indexToFlatRows", () => {
    it("emits one row per (canonicalKey × backlink) sorted by key", () => {
      const agg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(1, "A", "/sites/hr/SitePages/A.aspx", [
            { url: "x", canonicalKey: "/sites/z/z.pdf" },
            { url: "y", canonicalKey: "/sites/a/a.pdf" },
          ]),
          mkPage(2, "B", "/sites/hr/SitePages/B.aspx", [
            { url: "x", canonicalKey: "/sites/a/a.pdf" },
          ]),
        ]),
      ];
      const idx = buildBacklinksIndex(agg, "job-1", "2026-04-22T10:00:00Z");
      const rows = indexToFlatRows(idx);
      assert.equal(rows.length, 3);
      assert.equal(rows[0].canonicalKey, "/sites/a/a.pdf");
      assert.equal(rows[1].canonicalKey, "/sites/a/a.pdf");
      assert.equal(rows[2].canonicalKey, "/sites/z/z.pdf");
      assert.equal(rows[0].sourceUpdatedAt, "2026-04-22T10:00:00Z");
    });
  });
});
