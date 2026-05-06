import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import { extractBannerImageUrl, type SitePageItem } from "../services/spoPagesClient.js";
import { buildPageInventory } from "../services/linkInventoryScanner.js";

describe("extractBannerImageUrl", () => {
  it("returns undefined for missing field", () => {
    assert.equal(extractBannerImageUrl(undefined), undefined);
    assert.equal(extractBannerImageUrl(null), undefined);
  });

  it("returns undefined for empty string", () => {
    assert.equal(extractBannerImageUrl(""), undefined);
  });

  it("returns the string when the field is a plain string", () => {
    const url = "/sites/foo/SiteAssets/SitePages/X/banner.jpg";
    assert.equal(extractBannerImageUrl(url), url);
  });

  it("unwraps the Url property when SP returned a URL-field object", () => {
    const url = "https://contoso.sharepoint.com/sites/foo/SiteAssets/SitePages/X/banner.jpg";
    assert.equal(
      extractBannerImageUrl({ Url: url, Description: "/sites/foo/..." }),
      url,
    );
  });

  it("returns undefined for a URL-field object with empty/missing Url", () => {
    assert.equal(extractBannerImageUrl({ Url: "", Description: "x" }), undefined);
    assert.equal(extractBannerImageUrl({ Description: "x" }), undefined);
  });
});

describe("buildPageInventory: banner image synthesis", () => {
  it("emits a synthetic banner link when BannerImageUrl is set", () => {
    const item: SitePageItem = {
      Id: 1,
      FileLeafRef: "MyPage.aspx",
      FileRef: "/sites/foo/SitePages/MyPage.aspx",
      CanvasContent1: "",
      LayoutWebpartsContent: "",
      BannerImageUrl: {
        Url: "https://contoso.sharepoint.com/sites/foo/SiteAssets/SitePages/MyPage/banner.jpg",
        Description: "/sites/foo/SiteAssets/SitePages/MyPage/banner.jpg",
      },
    };

    const inv = buildPageInventory(item);
    const banners = inv.links.filter((l) => l.source === "banner");
    assert.equal(banners.length, 1);
    assert.equal(
      banners[0].url,
      "https://contoso.sharepoint.com/sites/foo/SiteAssets/SitePages/MyPage/banner.jpg",
    );
    // Must produce a non-empty canonical key so the backlinks index can index it
    assert.ok(banners[0].canonicalKey, "banner link should produce a canonical key");
    // Should classify as spo-internal (absolute URL on the tenant host).
    // Tenant host defaults to "contoso.sharepoint.com" via test-setup.ts.
    assert.equal(banners[0].linkClass, "spo-internal");
  });

  it("emits a banner link when BannerImageUrl is a plain string", () => {
    const item: SitePageItem = {
      Id: 2,
      FileLeafRef: "P.aspx",
      FileRef: "/sites/foo/SitePages/P.aspx",
      CanvasContent1: "",
      LayoutWebpartsContent: "",
      BannerImageUrl: "/sites/foo/SiteAssets/SitePages/P/hero.png",
    };

    const inv = buildPageInventory(item);
    const banners = inv.links.filter((l) => l.source === "banner");
    assert.equal(banners.length, 1);
    assert.equal(banners[0].url, "/sites/foo/SiteAssets/SitePages/P/hero.png");
    assert.equal(banners[0].linkClass, "relative");
  });

  it("does NOT emit a banner link when BannerImageUrl is empty/missing", () => {
    const item: SitePageItem = {
      Id: 3,
      FileLeafRef: "P.aspx",
      FileRef: "/sites/foo/SitePages/P.aspx",
      CanvasContent1: "",
      LayoutWebpartsContent: "",
      BannerImageUrl: null,
    };

    const inv = buildPageInventory(item);
    assert.equal(inv.links.filter((l) => l.source === "banner").length, 0);
  });

  it("banner link sits alongside canvas-extracted links", () => {
    // Realistic case: a page has both an inline anchor (caught by the
    // canvas extractor's RTE/anchor path) and a banner (caught by the
    // new banner synthesis).
    //
    // Use a Text web part's data-sp-rte region with an inline <a href>;
    // the anchor extractor handles that path with no JSON parsing, so
    // we don't need to forge entity-encoded controldata JSON to exercise
    // the dual-source case.
    const canvas =
      `<div data-sp-canvascontrol data-sp-controldata="{}">` +
      `<div data-sp-rte="">` +
      `<p>See <a href="/sites/foo/SiteAssets/SitePages/P/inline.png">the diagram</a></p>` +
      `</div>` +
      `</div>`;
    const item: SitePageItem = {
      Id: 4,
      FileLeafRef: "P.aspx",
      FileRef: "/sites/foo/SitePages/P.aspx",
      CanvasContent1: canvas,
      LayoutWebpartsContent: "",
      BannerImageUrl: "/sites/foo/SiteAssets/SitePages/P/banner.jpg",
    };

    const inv = buildPageInventory(item);
    const urls = inv.links.map((l) => l.url);
    assert.ok(
      urls.includes("/sites/foo/SiteAssets/SitePages/P/inline.png"),
      `expected inline.png in extracted links; got: ${JSON.stringify(urls)}`,
    );
    assert.ok(
      urls.includes("/sites/foo/SiteAssets/SitePages/P/banner.jpg"),
      `expected banner.jpg in extracted links; got: ${JSON.stringify(urls)}`,
    );
    // Exactly one of the links must be sourced as 'banner'
    assert.equal(inv.links.filter((l) => l.source === "banner").length, 1);
  });
});
