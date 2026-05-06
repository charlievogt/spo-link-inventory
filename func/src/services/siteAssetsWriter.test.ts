import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import {
  deriveSiteContextFromFile,
  parseRecycleResponse,
  recycleSitePagesAsset,
} from "./siteAssetsWriter.js";

describe("deriveSiteContextFromFile", () => {
  it("extracts /sites/<name> from a sites-collection URL", () => {
    assert.equal(
      deriveSiteContextFromFile("/sites/foo/SiteAssets/SitePages/p/img.png"),
      "/sites/foo",
    );
  });

  it("extracts /teams/<name> from a teams-collection URL", () => {
    assert.equal(
      deriveSiteContextFromFile("/teams/marketing/SiteAssets/SitePages/p/img.png"),
      "/teams/marketing",
    );
  });

  it("returns empty string for root-web URLs", () => {
    assert.equal(
      deriveSiteContextFromFile("/SiteAssets/SitePages/p/img.png"),
      "",
    );
  });

  it("returns empty string for unrecognized prefixes", () => {
    assert.equal(deriveSiteContextFromFile("/some/other/path.png"), "");
  });

  it("is case-insensitive on the sites/teams prefix", () => {
    assert.equal(
      deriveSiteContextFromFile("/Sites/Foo/SiteAssets/SitePages/p/img.png"),
      "/Sites/Foo",
    );
  });
});

describe("parseRecycleResponse", () => {
  it("pulls value from nometadata shape", () => {
    assert.equal(
      parseRecycleResponse({ value: "11111111-1111-1111-1111-111111111111" }),
      "11111111-1111-1111-1111-111111111111",
    );
  });

  it("pulls Recycle from minimalmetadata shape", () => {
    assert.equal(
      parseRecycleResponse({ Recycle: "22222222-2222-2222-2222-222222222222" }),
      "22222222-2222-2222-2222-222222222222",
    );
  });

  it("pulls d.Recycle from verbose shape", () => {
    assert.equal(
      parseRecycleResponse({ d: { Recycle: "33333333-3333-3333-3333-333333333333" } }),
      "33333333-3333-3333-3333-333333333333",
    );
  });

  it("returns undefined for unparseable shapes", () => {
    assert.equal(parseRecycleResponse(null), undefined);
    assert.equal(parseRecycleResponse(undefined), undefined);
    assert.equal(parseRecycleResponse({}), undefined);
    assert.equal(parseRecycleResponse({ d: {} }), undefined);
    assert.equal(parseRecycleResponse({ d: { Recycle: 42 } }), undefined);
    assert.equal(parseRecycleResponse("string"), undefined);
  });
});

describe("recycleSitePagesAsset", () => {
  function fakeFetch(
    response: Response | Error | (() => Response | Promise<Response>),
  ): typeof fetch {
    return async (input, init) => {
      // Mark unused-arg lint pacification
      void input;
      void init;
      if (response instanceof Error) throw response;
      if (typeof response === "function") {
        const r = response();
        return r instanceof Promise ? r : r;
      }
      return response;
    };
  }

  const url = "/sites/foo/SiteAssets/SitePages/X/orphan.png";

  it("returns recycled with the recycle bin item ID on 200", async () => {
    const fetchImpl = fakeFetch(
      new Response(JSON.stringify({ value: "abc-123" }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      }),
    );
    const result = await recycleSitePagesAsset(url, "fake-token", { fetchImpl });
    assert.equal(result.status, "recycled");
    assert.equal(result.recycleBinItemId, "abc-123");
    assert.equal(result.serverRelativeUrl, url);
  });

  it("still returns recycled when the body is unparseable", async () => {
    const fetchImpl = fakeFetch(
      new Response("not json", {
        status: 200,
        headers: { "Content-Type": "text/plain" },
      }),
    );
    const result = await recycleSitePagesAsset(url, "fake-token", { fetchImpl });
    assert.equal(result.status, "recycled");
    assert.equal(result.recycleBinItemId, undefined);
  });

  it("returns not-found for 404", async () => {
    const fetchImpl = fakeFetch(
      new Response("File Not Found", { status: 404 }),
    );
    const result = await recycleSitePagesAsset(url, "fake-token", { fetchImpl });
    assert.equal(result.status, "not-found");
    assert.match(result.error ?? "", /File Not Found/);
  });

  it("returns forbidden for 403", async () => {
    const fetchImpl = fakeFetch(
      new Response("Access denied", { status: 403 }),
    );
    const result = await recycleSitePagesAsset(url, "fake-token", { fetchImpl });
    assert.equal(result.status, "forbidden");
    assert.match(result.error ?? "", /delete permission/i);
  });

  it("returns error for other non-200s (e.g. 500, throttle)", async () => {
    const fetchImpl = fakeFetch(
      new Response("throttled", { status: 429 }),
    );
    const result = await recycleSitePagesAsset(url, "fake-token", { fetchImpl });
    assert.equal(result.status, "error");
    assert.match(result.error ?? "", /429/);
  });

  it("returns error when fetch throws (network failure)", async () => {
    const fetchImpl = fakeFetch(new TypeError("fetch failed"));
    const result = await recycleSitePagesAsset(url, "fake-token", { fetchImpl });
    assert.equal(result.status, "error");
    assert.match(result.error ?? "", /fetch threw.*fetch failed/);
  });

  it("constructs the correct SP REST URL: site context concatenated raw, file path per-segment encoded", async () => {
    // Note: site context is concatenated raw (relying on fetch/URL to
    // normalize literal spaces — matches the pattern in
    // linkInventoryWriter.patchPageCanvas and spoFilesEnumerator.fetchFileBytes).
    // The file path inside getfilebyserverrelativeurl(...) is per-segment
    // encoded so spaces become %20 in the OData filter.
    let observedUrl: string | undefined;
    let observedAuth: string | undefined;
    const fetchImpl: typeof fetch = async (input, init) => {
      observedUrl = typeof input === "string" ? input : input.toString();
      observedAuth = ((init?.headers as Record<string, string>) ?? {}).Authorization;
      return new Response(JSON.stringify({ value: "ok" }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    };

    await recycleSitePagesAsset(
      "/sites/Marketing Hub/SiteAssets/SitePages/Q1 Plan/diagram.png",
      "the-token",
      { fetchImpl },
    );

    assert.ok(observedUrl);
    // Site context segment is concatenated raw (fetch/URL will encode it)
    assert.match(observedUrl!, /\/sites\/Marketing Hub\/_api\/web\/getfilebyserverrelativeurl/);
    // The full path in the OData filter is per-segment encoded
    assert.match(
      observedUrl!,
      /getfilebyserverrelativeurl\('\/sites\/Marketing%20Hub\/SiteAssets\/SitePages\/Q1%20Plan\/diagram\.png'\)\/recycle\(\)/,
    );
    assert.equal(observedAuth, "Bearer the-token");
  });

  it("doubles single quotes (SP OData escape) in path segments", async () => {
    let observedUrl: string | undefined;
    const fetchImpl: typeof fetch = async (input) => {
      observedUrl = typeof input === "string" ? input : input.toString();
      return new Response(JSON.stringify({ value: "ok" }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    };
    await recycleSitePagesAsset(
      "/sites/foo/SiteAssets/SitePages/O'Brien/img.png",
      "tok",
      { fetchImpl },
    );
    assert.match(observedUrl ?? "", /O''Brien/);
  });
});
