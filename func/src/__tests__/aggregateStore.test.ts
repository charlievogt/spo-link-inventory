import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import {
  emptySiteAggregate,
  mergePageScanIntoAggregate,
  mergeDocScanIntoAggregate,
  carryForwardEntries,
  siteSlug,
  type SiteAggregate,
  type PageScanResult,
  type FileScanResult,
} from "../services/aggregateStore.js";
import type { ClassifiedLink } from "../services/linkInventoryScanner.js";
import type { ClassifiedDocLink } from "../services/documentLinkScanner.js";

function mkPageLink(url: string): ClassifiedLink {
  return {
    rawUrl: url,
    url,
    source: "anchor",
    linkClass: "spo-internal",
    normalizedKey: url,
    canonicalKey: url,
  };
}

function mkDocLink(url: string): ClassifiedDocLink {
  return {
    rawUrl: url,
    url,
    source: "ooxml-relationship",
    linkClass: "spo-internal",
    normalizedKey: "",
  };
}

describe("aggregateStore.siteSlug", () => {
  it("collapses /sites/<x> paths to slugs", () => {
    assert.equal(siteSlug("/sites/charlie-test-site"), "charlie-test-site");
    assert.equal(siteSlug("/sites/HR_Site"), "hr-site");
    assert.equal(siteSlug("/sites/sales/marketing"), "sales-marketing");
  });

  it("handles root web", () => {
    assert.equal(siteSlug("/"), "root");
  });

  it("strips leading/trailing dashes after substitution", () => {
    assert.equal(siteSlug("///"), "root");
  });
});

describe("aggregateStore.mergePageScanIntoAggregate", () => {
  const site = "/sites/hr";

  it("inserts new page entries", () => {
    const initial = emptySiteAggregate(site);
    const results: PageScanResult[] = [
      {
        pageUrl: "/sites/hr/SitePages/handbook.aspx",
        pageTitle: "Handbook",
        modified: "2026-04-25T10:00:00Z",
        links: [mkPageLink("https://contoso.sharepoint.com/sites/hr")],
      },
    ];
    const next = mergePageScanIntoAggregate(initial, results, "job-1", "2026-04-25T11:00:00Z");
    assert.equal(Object.keys(next.pages).length, 1);
    const e = next.pages["/sites/hr/SitePages/handbook.aspx"];
    assert.equal(e.pageTitle, "Handbook");
    assert.equal(e.lastConfirmedByJobId, "job-1");
    assert.equal(e.lastConfirmedAt, "2026-04-25T11:00:00Z");
    assert.equal(next.lastPageScanAt, "2026-04-25T11:00:00Z");
  });

  it("replaces existing entries on re-scan", () => {
    const initial: SiteAggregate = {
      ...emptySiteAggregate(site),
      pages: {
        "/sites/hr/SitePages/handbook.aspx": {
          pageUrl: "/sites/hr/SitePages/handbook.aspx",
          pageTitle: "Handbook",
          modified: "2026-04-01T10:00:00Z",
          links: [mkPageLink("https://old-link")],
          lastConfirmedAt: "2026-04-01T11:00:00Z",
          lastConfirmedByJobId: "job-old",
        },
      },
    };
    const results: PageScanResult[] = [
      {
        pageUrl: "/sites/hr/SitePages/handbook.aspx",
        pageTitle: "Handbook v2",
        modified: "2026-04-25T10:00:00Z",
        links: [mkPageLink("https://new-link")],
      },
    ];
    const next = mergePageScanIntoAggregate(initial, results, "job-2", "2026-04-25T11:00:00Z");
    const e = next.pages["/sites/hr/SitePages/handbook.aspx"];
    assert.equal(e.pageTitle, "Handbook v2");
    assert.equal(e.modified, "2026-04-25T10:00:00Z");
    assert.equal(e.lastConfirmedByJobId, "job-2");
    assert.deepEqual(e.links, [mkPageLink("https://new-link")]);
  });

  it("leaves pages absent from results untouched", () => {
    const initial: SiteAggregate = {
      ...emptySiteAggregate(site),
      pages: {
        "/sites/hr/SitePages/handbook.aspx": {
          pageUrl: "/sites/hr/SitePages/handbook.aspx",
          pageTitle: "Handbook",
          modified: "2026-04-01T10:00:00Z",
          links: [mkPageLink("https://kept")],
          lastConfirmedAt: "2026-04-01T11:00:00Z",
          lastConfirmedByJobId: "job-old",
        },
      },
    };
    const next = mergePageScanIntoAggregate(initial, [], "job-2", "2026-04-25T11:00:00Z");
    const e = next.pages["/sites/hr/SitePages/handbook.aspx"];
    assert.equal(e.lastConfirmedByJobId, "job-old", "untouched entry retains old jobId");
    assert.deepEqual(e.links, [mkPageLink("https://kept")]);
  });
});

describe("aggregateStore.mergeDocScanIntoAggregate", () => {
  const site = "/sites/hr";

  it("inserts new file entries with all metadata", () => {
    const initial = emptySiteAggregate(site);
    const results: FileScanResult[] = [
      {
        fileRef: "/sites/hr/Shared Documents/policy.pdf",
        fileName: "policy.pdf",
        fileType: "pdf",
        modified: "2026-04-25T10:00:00Z",
        etag: '"{abc},5"',
        size: 24576,
        links: [mkDocLink("https://intranet/about")],
      },
    ];
    const next = mergeDocScanIntoAggregate(initial, results, "job-1", "2026-04-25T11:00:00Z");
    const e = next.files["/sites/hr/Shared Documents/policy.pdf"];
    assert.equal(e.fileType, "pdf");
    assert.equal(e.etag, '"{abc},5"');
    assert.equal(e.size, 24576);
    assert.equal(e.lastConfirmedAt, "2026-04-25T11:00:00Z");
    assert.equal(next.lastDocScanAt, "2026-04-25T11:00:00Z");
  });

  it("clears scanError when re-scan succeeds", () => {
    const initial: SiteAggregate = {
      ...emptySiteAggregate(site),
      files: {
        "/sites/hr/Shared Documents/x.pdf": {
          fileRef: "/sites/hr/Shared Documents/x.pdf",
          fileName: "x.pdf",
          fileType: "pdf",
          modified: "2026-04-01T10:00:00Z",
          etag: '"{old},1"',
          size: 1024,
          links: [],
          scanError: "download failed: 503",
          lastConfirmedAt: "2026-04-01T11:00:00Z",
          lastConfirmedByJobId: "job-old",
        },
      },
    };
    const results: FileScanResult[] = [
      {
        fileRef: "/sites/hr/Shared Documents/x.pdf",
        fileName: "x.pdf",
        fileType: "pdf",
        modified: "2026-04-25T10:00:00Z",
        etag: '"{new},2"',
        size: 1024,
        links: [mkDocLink("https://example/a")],
        // no scanError on success
      },
    ];
    const next = mergeDocScanIntoAggregate(initial, results, "job-2", "2026-04-25T11:00:00Z");
    const e = next.files["/sites/hr/Shared Documents/x.pdf"];
    assert.equal(e.scanError, undefined);
    assert.equal(e.etag, '"{new},2"');
    assert.equal(e.links.length, 1);
  });

  it("leaves files absent from results untouched", () => {
    const initial: SiteAggregate = {
      ...emptySiteAggregate(site),
      files: {
        "/sites/hr/Shared Documents/kept.pdf": {
          fileRef: "/sites/hr/Shared Documents/kept.pdf",
          fileName: "kept.pdf",
          fileType: "pdf",
          modified: "2026-04-01T10:00:00Z",
          etag: '"{kept},1"',
          size: 1024,
          links: [mkDocLink("https://kept")],
          lastConfirmedAt: "2026-04-01T11:00:00Z",
          lastConfirmedByJobId: "job-old",
        },
      },
    };
    const next = mergeDocScanIntoAggregate(initial, [], "job-2", "2026-04-25T11:00:00Z");
    const e = next.files["/sites/hr/Shared Documents/kept.pdf"];
    assert.equal(e.lastConfirmedByJobId, "job-old");
    assert.equal(e.etag, '"{kept},1"');
  });
});

describe("aggregateStore.carryForwardEntries", () => {
  const site = "/sites/hr";

  it("bumps lastConfirmedAt on existing pages without touching content", () => {
    const initial: SiteAggregate = {
      ...emptySiteAggregate(site),
      pages: {
        "/sites/hr/SitePages/handbook.aspx": {
          pageUrl: "/sites/hr/SitePages/handbook.aspx",
          pageTitle: "Handbook",
          modified: "2026-04-01T10:00:00Z",
          links: [mkPageLink("https://kept")],
          lastConfirmedAt: "2026-04-01T11:00:00Z",
          lastConfirmedByJobId: "job-old",
        },
      },
    };
    const next = carryForwardEntries(
      initial,
      ["/sites/hr/SitePages/handbook.aspx"],
      [],
      "job-2",
      "2026-04-25T11:00:00Z",
    );
    const e = next.pages["/sites/hr/SitePages/handbook.aspx"];
    assert.equal(e.lastConfirmedAt, "2026-04-25T11:00:00Z");
    assert.equal(e.lastConfirmedByJobId, "job-2");
    assert.equal(e.modified, "2026-04-01T10:00:00Z", "modified untouched");
    assert.deepEqual(e.links, [mkPageLink("https://kept")], "links untouched");
  });

  it("bumps lastConfirmedAt on existing files without touching content", () => {
    const initial: SiteAggregate = {
      ...emptySiteAggregate(site),
      files: {
        "/sites/hr/Shared Documents/policy.pdf": {
          fileRef: "/sites/hr/Shared Documents/policy.pdf",
          fileName: "policy.pdf",
          fileType: "pdf",
          modified: "2026-04-01T10:00:00Z",
          etag: '"{abc},1"',
          size: 24576,
          links: [mkDocLink("https://kept")],
          lastConfirmedAt: "2026-04-01T11:00:00Z",
          lastConfirmedByJobId: "job-old",
        },
      },
    };
    const next = carryForwardEntries(
      initial,
      [],
      ["/sites/hr/Shared Documents/policy.pdf"],
      "job-2",
      "2026-04-25T11:00:00Z",
    );
    const e = next.files["/sites/hr/Shared Documents/policy.pdf"];
    assert.equal(e.lastConfirmedAt, "2026-04-25T11:00:00Z");
    assert.equal(e.lastConfirmedByJobId, "job-2");
    assert.equal(e.etag, '"{abc},1"', "etag untouched");
  });

  it("silently ignores keys that aren't in the aggregate", () => {
    const initial = emptySiteAggregate(site);
    const next = carryForwardEntries(
      initial,
      ["/sites/hr/SitePages/never-seen.aspx"],
      ["/sites/hr/Shared Documents/never-seen.pdf"],
      "job-2",
      "2026-04-25T11:00:00Z",
    );
    assert.equal(Object.keys(next.pages).length, 0);
    assert.equal(Object.keys(next.files).length, 0);
  });
});

describe("aggregateStore — ETag-skip workflow integration", () => {
  // End-to-end shape check: simulate "first scan populates aggregate,
  // second scan delta-skips one file → carryForward, modifies another →
  // mergeDocScanIntoAggregate". Verifies the merge primitives compose
  // correctly without dropping or double-counting entries.
  it("first scan creates entries; second scan with one delta + one skip rolls forward correctly", () => {
    const site = "/sites/hr";

    // First scan populates two files
    const initialResults: FileScanResult[] = [
      {
        fileRef: "/sites/hr/docs/a.pdf",
        fileName: "a.pdf",
        fileType: "pdf",
        modified: "2026-04-01T10:00:00Z",
        etag: '"{a},1"',
        size: 100,
        links: [mkDocLink("https://link-a")],
      },
      {
        fileRef: "/sites/hr/docs/b.pdf",
        fileName: "b.pdf",
        fileType: "pdf",
        modified: "2026-04-01T10:00:00Z",
        etag: '"{b},1"',
        size: 200,
        links: [mkDocLink("https://link-b")],
      },
    ];
    let agg = emptySiteAggregate(site);
    agg = mergeDocScanIntoAggregate(agg, initialResults, "job-1", "2026-04-01T11:00:00Z");

    // Second scan: a.pdf changed (etag mismatch — full re-scan), b.pdf
    // unchanged (etag match — carryForward).
    const secondResults: FileScanResult[] = [
      {
        fileRef: "/sites/hr/docs/a.pdf",
        fileName: "a.pdf",
        fileType: "pdf",
        modified: "2026-04-25T10:00:00Z",
        etag: '"{a},2"',
        size: 150,
        links: [mkDocLink("https://link-a-updated")],
      },
    ];
    agg = mergeDocScanIntoAggregate(agg, secondResults, "job-2", "2026-04-25T11:00:00Z");
    agg = carryForwardEntries(agg, [], ["/sites/hr/docs/b.pdf"], "job-2", "2026-04-25T11:00:00Z");

    // a.pdf should reflect new content
    assert.equal(agg.files["/sites/hr/docs/a.pdf"].etag, '"{a},2"');
    assert.equal(agg.files["/sites/hr/docs/a.pdf"].size, 150);
    assert.deepEqual(
      agg.files["/sites/hr/docs/a.pdf"].links,
      [mkDocLink("https://link-a-updated")],
    );
    assert.equal(agg.files["/sites/hr/docs/a.pdf"].lastConfirmedByJobId, "job-2");

    // b.pdf should retain original content but updated lastConfirmedAt
    assert.equal(agg.files["/sites/hr/docs/b.pdf"].etag, '"{b},1"', "skipped etag preserved");
    assert.deepEqual(
      agg.files["/sites/hr/docs/b.pdf"].links,
      [mkDocLink("https://link-b")],
      "skipped links preserved",
    );
    assert.equal(
      agg.files["/sites/hr/docs/b.pdf"].lastConfirmedByJobId,
      "job-2",
      "skipped entry's confirmation jobId bumped",
    );
    assert.equal(
      agg.files["/sites/hr/docs/b.pdf"].lastConfirmedAt,
      "2026-04-25T11:00:00Z",
    );
  });
});
