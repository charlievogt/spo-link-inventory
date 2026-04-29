import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import {
  buildSectionSlicesFromScan,
  mergeScanIntoIndex,
  buildPersistentFromAggregates,
  inflateToMemory,
  type PersistentBacklinksIndex,
} from "../services/backlinksIndexStore.js";
import type { ISiteInventoryShape } from "../functions/linkInventoryReplace.types.js";

function mkPage(
  pageId: number,
  pageTitle: string,
  pageUrl: string,
  links: Array<{ canonicalKey?: string }>,
): ISiteInventoryShape["pages"][number] {
  return {
    pageId,
    pageTitle,
    pageUrl,
    links: links.map((l) => ({
      rawUrl: "x",
      url: "x",
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
  return { site, pageCount: pages.length, linkCount: 0, pages };
}

function emptyPersistent(): PersistentBacklinksIndex {
  return { version: 1, builtAt: "2026-01-01T00:00:00Z", bySite: {} };
}

describe("backlinksIndexStore", () => {
  describe("buildSectionSlicesFromScan", () => {
    it("produces one slice per site covered by the scan", () => {
      const agg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(1, "Benefits", "/sites/hr/SitePages/Benefits.aspx", [{ canonicalKey: "/sites/x/foo.pdf" }]),
        ]),
        mkSite("/sites/finance", [
          mkPage(1, "Budget", "/sites/finance/SitePages/Budget.aspx", [{ canonicalKey: "/sites/x/bar.pdf" }]),
        ]),
      ];
      const slices = buildSectionSlicesFromScan(agg, "job-1", "2026-04-22T10:00:00Z");
      assert.equal(slices.size, 2);
      assert.ok(slices.has("/sites/hr"));
      assert.ok(slices.has("/sites/finance"));
    });

    it("lowercases site paths for the key", () => {
      const agg: ISiteInventoryShape[] = [
        mkSite("/Sites/HR", [mkPage(1, "X", "/x", [{ canonicalKey: "/y" }])]),
      ];
      const slices = buildSectionSlicesFromScan(agg, "j", "t");
      assert.ok(slices.has("/sites/hr"));
    });

    it("skips sources with no outbound links but keeps the site entry", () => {
      const agg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(1, "External-only", "/sites/hr/x", [{ canonicalKey: undefined }]),
        ]),
      ];
      const slices = buildSectionSlicesFromScan(agg, "j", "t");
      const s = slices.get("/sites/hr");
      assert.ok(s);
      assert.equal(s.sources.length, 0, "no sources pass the link filter");
    });

    it("dedupes the same canonicalKey within one source", () => {
      const agg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [
          mkPage(1, "Multi", "/sites/hr/p", [
            { canonicalKey: "/sites/x/foo.pdf" },
            { canonicalKey: "/sites/x/foo.pdf" },
            { canonicalKey: "contoso.sharepoint.com/sites/x/foo.pdf" },
          ]),
        ]),
      ];
      const slices = buildSectionSlicesFromScan(agg, "j", "t");
      const s = slices.get("/sites/hr");
      assert.ok(s);
      assert.equal(s.sources.length, 1);
      assert.deepEqual(s.sources[0].linksTo, ["/sites/x/foo.pdf"]);
    });
  });

  describe("mergeScanIntoIndex", () => {
    it("adds a fresh pages slice to an empty index", () => {
      const slices = buildSectionSlicesFromScan(
        [mkSite("/sites/hr", [mkPage(1, "A", "/a", [{ canonicalKey: "/t" }])])],
        "job-1",
        "2026-04-22T10:00:00Z",
      );
      const merged = mergeScanIntoIndex(emptyPersistent(), slices, "page", ["/sites/hr"]);
      assert.ok(merged.bySite["/sites/hr"]);
      assert.ok(merged.bySite["/sites/hr"].pages);
      assert.equal(merged.bySite["/sites/hr"].pages?.fromJobId, "job-1");
      assert.equal(merged.bySite["/sites/hr"].documents, undefined);
    });

    it("replaces only the kind-specific slice, preserving the other kind", () => {
      // Seed: /sites/hr has pages + documents from old jobs
      let idx: PersistentBacklinksIndex = {
        version: 1,
        builtAt: "2026-01-01T00:00:00Z",
        bySite: {
          "/sites/hr": {
            siteLabel: "hr",
            siteUrl: "https://contoso.sharepoint.com/sites/hr",
            pages: { updatedAt: "2026-03-01T00:00:00Z", fromJobId: "old-page-job", sources: [] },
            documents: { updatedAt: "2026-03-05T00:00:00Z", fromJobId: "old-doc-job", sources: [] },
          },
        },
      };

      const newPageSlices = buildSectionSlicesFromScan(
        [mkSite("/sites/hr", [mkPage(1, "NewPage", "/sites/hr/new", [{ canonicalKey: "/t" }])])],
        "new-page-job",
        "2026-04-22T10:00:00Z",
      );
      idx = mergeScanIntoIndex(idx, newPageSlices, "page", ["/sites/hr"]);

      const hr = idx.bySite["/sites/hr"];
      assert.equal(hr.pages?.fromJobId, "new-page-job", "pages slice replaced");
      assert.equal(hr.pages?.updatedAt, "2026-04-22T10:00:00Z");
      assert.equal(hr.documents?.fromJobId, "old-doc-job", "documents slice preserved");
      assert.equal(hr.documents?.updatedAt, "2026-03-05T00:00:00Z");
    });

    it("leaves sites not in scannedSitePaths untouched", () => {
      let idx: PersistentBacklinksIndex = {
        version: 1,
        builtAt: "2026-01-01T00:00:00Z",
        bySite: {
          "/sites/hr": {
            siteLabel: "hr",
            siteUrl: "https://contoso.sharepoint.com/sites/hr",
            pages: { updatedAt: "2026-03-01T00:00:00Z", fromJobId: "old", sources: [] },
          },
          "/sites/finance": {
            siteLabel: "finance",
            siteUrl: "https://contoso.sharepoint.com/sites/finance",
            pages: { updatedAt: "2026-03-01T00:00:00Z", fromJobId: "old", sources: [] },
          },
        },
      };

      const slices = buildSectionSlicesFromScan(
        [mkSite("/sites/hr", [mkPage(1, "A", "/a", [{ canonicalKey: "/t" }])])],
        "new-job",
        "2026-04-22T10:00:00Z",
      );
      idx = mergeScanIntoIndex(idx, slices, "page", ["/sites/hr"]);

      assert.equal(idx.bySite["/sites/hr"].pages?.fromJobId, "new-job");
      assert.equal(idx.bySite["/sites/finance"].pages?.fromJobId, "old", "finance untouched");
    });

    it("writes an empty pages slice for a site that was scanned with zero sources", () => {
      // Scan covered /sites/empty but produced no sources (no canonicalKeys).
      const slices = buildSectionSlicesFromScan(
        [mkSite("/sites/empty", [mkPage(1, "Ext-only", "/p", [{ canonicalKey: undefined }])])],
        "job-2",
        "2026-04-22T10:00:00Z",
      );
      const idx = mergeScanIntoIndex(emptyPersistent(), slices, "page", ["/sites/empty"]);
      const s = idx.bySite["/sites/empty"];
      assert.ok(s);
      assert.ok(s.pages, "pages slice exists");
      assert.equal(s.pages?.sources.length, 0);
      assert.equal(s.pages?.fromJobId, "job-2");
    });

    it("doc scan merges into documents slice without touching pages", () => {
      let idx: PersistentBacklinksIndex = {
        version: 1,
        builtAt: "2026-01-01T00:00:00Z",
        bySite: {
          "/sites/hr": {
            siteLabel: "hr",
            siteUrl: "https://contoso.sharepoint.com/sites/hr",
            pages: { updatedAt: "2026-03-01T00:00:00Z", fromJobId: "page-job", sources: [] },
          },
        },
      };
      // Doc scan finalize wraps files into page-shaped records — that's
      // what mkPage/mkSite produces here.
      const slices = buildSectionSlicesFromScan(
        [mkSite("/sites/hr", [mkPage(0, "report.docx", "/sites/hr/Shared Documents/report.docx", [{ canonicalKey: "/t" }])])],
        "doc-job",
        "2026-04-22T10:00:00Z",
      );
      idx = mergeScanIntoIndex(idx, slices, "document", ["/sites/hr"]);

      assert.equal(idx.bySite["/sites/hr"].pages?.fromJobId, "page-job", "pages preserved");
      assert.equal(idx.bySite["/sites/hr"].documents?.fromJobId, "doc-job");
      assert.equal(idx.bySite["/sites/hr"].documents?.sources[0].title, "report.docx");
    });
  });

  describe("buildPersistentFromAggregates (rebuild)", () => {
    it("combines page + doc aggregates into one persistent index", () => {
      const pageAgg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [mkPage(1, "Benefits", "/b", [{ canonicalKey: "/t1" }])]),
      ];
      const docAgg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [mkPage(0, "procedure.docx", "/pd", [{ canonicalKey: "/t2" }])]),
        mkSite("/sites/finance", [mkPage(0, "budget.xlsx", "/bg", [{ canonicalKey: "/t3" }])]),
      ];
      const idx = buildPersistentFromAggregates(
        pageAgg, "page-job", "2026-04-22T10:00:00Z",
        docAgg, "doc-job", "2026-04-22T11:00:00Z",
      );
      assert.ok(idx.bySite["/sites/hr"].pages);
      assert.ok(idx.bySite["/sites/hr"].documents);
      assert.ok(idx.bySite["/sites/finance"].documents);
      assert.equal(idx.bySite["/sites/finance"].pages, undefined);
    });

    it("handles null aggregates (only one kind available)", () => {
      const pageAgg: ISiteInventoryShape[] = [
        mkSite("/sites/hr", [mkPage(1, "A", "/a", [{ canonicalKey: "/t" }])]),
      ];
      const idx = buildPersistentFromAggregates(
        pageAgg, "pj", "2026-04-22T10:00:00Z",
        null, null, null,
      );
      assert.ok(idx.bySite["/sites/hr"].pages);
      assert.equal(idx.bySite["/sites/hr"].documents, undefined);
    });
  });

  describe("inflateToMemory", () => {
    it("inverts the persistent index into a canonicalKey lookup map", () => {
      const persistent: PersistentBacklinksIndex = {
        version: 1,
        builtAt: "2026-04-22T10:00:00Z",
        bySite: {
          "/sites/hr": {
            siteLabel: "hr",
            siteUrl: "https://contoso.sharepoint.com/sites/hr",
            pages: {
              updatedAt: "2026-04-22T10:00:00Z",
              fromJobId: "pj",
              sources: [
                {
                  title: "Benefits",
                  url: "https://contoso.sharepoint.com/sites/hr/SitePages/Benefits.aspx",
                  linksTo: ["/sites/x/foo.pdf", "/sites/x/bar.pdf"],
                },
              ],
            },
            documents: {
              updatedAt: "2026-04-23T10:00:00Z",
              fromJobId: "dj",
              sources: [
                {
                  title: "manual.docx",
                  url: "https://contoso.sharepoint.com/sites/hr/Shared Documents/manual.docx",
                  linksTo: ["/sites/x/foo.pdf"],
                },
              ],
            },
          },
        },
      };
      const idx = inflateToMemory(persistent);
      const foo = idx.byCanonicalKey.get("/sites/x/foo.pdf");
      assert.ok(foo);
      assert.equal(foo.length, 2, "foo.pdf is linked from 1 page + 1 doc");
      const kinds = foo.map((e) => e.sourceKind).sort();
      assert.deepEqual(kinds, ["document", "page"]);

      const bar = idx.byCanonicalKey.get("/sites/x/bar.pdf");
      assert.ok(bar);
      assert.equal(bar.length, 1);
      assert.equal(bar[0].sourceKind, "page");

      // Top-level scannedAt is the max updatedAt across slices.
      assert.equal(idx.scannedAt, "2026-04-23T10:00:00Z");
    });

    it("per-entry sourceUpdatedAt comes from the entry's slice", () => {
      const persistent: PersistentBacklinksIndex = {
        version: 1,
        builtAt: "2026-04-22T10:00:00Z",
        bySite: {
          "/sites/hr": {
            siteLabel: "hr",
            siteUrl: "https://contoso.sharepoint.com/sites/hr",
            pages: {
              updatedAt: "2026-03-01T00:00:00Z",
              fromJobId: "pj",
              sources: [{ title: "A", url: "/a", linksTo: ["/t"] }],
            },
            documents: {
              updatedAt: "2026-04-22T10:00:00Z",
              fromJobId: "dj",
              sources: [{ title: "B", url: "/b", linksTo: ["/t"] }],
            },
          },
        },
      };
      const idx = inflateToMemory(persistent);
      const entries = idx.byCanonicalKey.get("/t");
      assert.ok(entries);
      const byKind = Object.fromEntries(entries.map((e) => [e.sourceKind, e.sourceUpdatedAt]));
      assert.equal(byKind.page, "2026-03-01T00:00:00Z");
      assert.equal(byKind.document, "2026-04-22T10:00:00Z");
    });
  });
});
