import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import {
  encodeServerRelativePath,
  parseFolderListing,
  enumerateSitePagesAssets,
} from "./spoSitePagesAssetsEnumerator.js";

describe("encodeServerRelativePath", () => {
  it("preserves slashes between segments", () => {
    assert.equal(
      encodeServerRelativePath("/sites/foo/SiteAssets/SitePages"),
      "/sites/foo/SiteAssets/SitePages",
    );
  });

  it("URL-encodes spaces and other unsafe chars per segment", () => {
    assert.equal(
      encodeServerRelativePath("/sites/Marketing Hub/SiteAssets/SitePages/Q1 Plan"),
      "/sites/Marketing%20Hub/SiteAssets/SitePages/Q1%20Plan",
    );
  });

  it("doubles single quotes (SP OData escape)", () => {
    assert.equal(
      encodeServerRelativePath("/sites/foo/SiteAssets/SitePages/O'Brien"),
      "/sites/foo/SiteAssets/SitePages/O''Brien",
    );
  });
});

describe("parseFolderListing", () => {
  const sitePath = "/sites/foo";

  it("handles the nometadata shape (flat arrays)", () => {
    const data = {
      Folders: [
        { ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/Page1", Name: "Page1" },
      ],
      Files: [
        {
          ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/orphan.png",
          Name: "orphan.png",
          Length: "12345",
          TimeLastModified: "2026-05-01T12:00:00Z",
        },
      ],
    };
    const out = parseFolderListing(data, sitePath);
    assert.equal(out.files.length, 1);
    assert.equal(out.files[0].size, 12345);
    assert.equal(out.files[0].name, "orphan.png");
    assert.equal(out.subfolders.length, 1);
    assert.equal(out.subfolders[0], "/sites/foo/SiteAssets/SitePages/Page1");
  });

  it("handles the minimalmetadata shape (.results wrapper)", () => {
    const data = {
      Folders: { results: [{ ServerRelativeUrl: "/x/y", Name: "y" }] },
      Files: {
        results: [
          {
            ServerRelativeUrl: "/x/y/a.png",
            Name: "a.png",
            Length: 100,
            TimeLastModified: "2026-01-01T00:00:00Z",
          },
        ],
      },
    };
    const out = parseFolderListing(data, sitePath);
    assert.equal(out.files.length, 1);
    assert.equal(out.subfolders.length, 1);
  });

  it("returns empty when input is null/undefined/non-object", () => {
    for (const bad of [null, undefined, "string", 42, true]) {
      const out = parseFolderListing(bad as unknown, sitePath);
      assert.deepEqual(out, { files: [], subfolders: [] });
    }
  });

  it("skips files with missing ServerRelativeUrl or Name", () => {
    const data = {
      Files: [
        { ServerRelativeUrl: "/x/y/good.png", Name: "good.png", Length: 1 },
        { ServerRelativeUrl: "/x/y/bad.png" }, // missing Name
        { Name: "alsobad.png", Length: 1 }, // missing ServerRelativeUrl
        null,
        "string",
      ],
    };
    const out = parseFolderListing(data, sitePath);
    assert.equal(out.files.length, 1);
    assert.equal(out.files[0].name, "good.png");
  });

  it("coerces numeric Length to a number", () => {
    const data = {
      Files: [
        { ServerRelativeUrl: "/x/a.png", Name: "a.png", Length: "9999" },
        { ServerRelativeUrl: "/x/b.png", Name: "b.png", Length: 8888 },
        { ServerRelativeUrl: "/x/c.png", Name: "c.png" }, // missing Length
      ],
    };
    const out = parseFolderListing(data, sitePath);
    assert.equal(out.files[0].size, 9999);
    assert.equal(out.files[1].size, 8888);
    assert.equal(out.files[2].size, 0);
  });

  it("skips underscore-prefixed system subfolders", () => {
    const data = {
      Folders: [
        { ServerRelativeUrl: "/x/Page1", Name: "Page1" },
        { ServerRelativeUrl: "/x/_catalogs", Name: "_catalogs" },
        { ServerRelativeUrl: "/x/Page2", Name: "Page2" },
      ],
    };
    const out = parseFolderListing(data, sitePath);
    assert.deepEqual(out.subfolders, ["/x/Page1", "/x/Page2"]);
  });
});

describe("enumerateSitePagesAssets", () => {
  function mockFetch(
    routes: Record<string, { status: number; body: unknown } | { status: number; bodyText: string }>,
  ): typeof fetch {
    return async (input: Parameters<typeof fetch>[0]): Promise<Response> => {
      const url = typeof input === "string" ? input : input.toString();
      // Match by the encoded folder path embedded in the URL — finds
      // the route entry whose key appears as a substring (so callers can
      // key by folder path without needing the absolute URL).
      const matched = Object.entries(routes).find(([key]) => url.includes(encodeURIComponent(key)) || url.includes(key));
      if (!matched) {
        return new Response("not stubbed", { status: 599, statusText: "no stub" });
      }
      const r = matched[1];
      if ("body" in r) {
        return new Response(JSON.stringify(r.body), {
          status: r.status,
          headers: { "Content-Type": "application/json" },
        });
      }
      return new Response(r.bodyText, { status: r.status });
    };
  }

  const tokenProvider = async (): Promise<string> => "fake-token";

  it("returns empty array when the root SitePages folder is missing (404)", async () => {
    const fetchImpl = mockFetch({
      "/SiteAssets/SitePages": { status: 404, bodyText: "File Not Found" },
    });
    const out = await enumerateSitePagesAssets("/sites/foo", { fetchImpl, tokenProvider });
    assert.deepEqual(out, []);
  });

  it("returns empty array when SP returns 400 'does not exist' for the root", async () => {
    const fetchImpl = mockFetch({
      "/SiteAssets/SitePages": { status: 400, bodyText: "Folder does not exist" },
    });
    const out = await enumerateSitePagesAssets("/sites/foo", { fetchImpl, tokenProvider });
    assert.deepEqual(out, []);
  });

  it("walks one level: returns files in the root folder", async () => {
    const fetchImpl = mockFetch({
      "/SiteAssets/SitePages": {
        status: 200,
        body: {
          Folders: [],
          Files: [
            {
              ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/stray.png",
              Name: "stray.png",
              Length: 100,
              TimeLastModified: "2026-01-01T00:00:00Z",
            },
          ],
        },
      },
    });
    const out = await enumerateSitePagesAssets("/sites/foo", { fetchImpl, tokenProvider });
    assert.equal(out.length, 1);
    assert.equal(out[0].name, "stray.png");
    assert.equal(out[0].sitePath, "/sites/foo");
  });

  it("recurses into subfolders", async () => {
    let calls = 0;
    const fetchImpl: typeof fetch = async (input) => {
      calls++;
      const url = typeof input === "string" ? input : input.toString();
      // The URL contains the folder path with slashes preserved (per-segment
      // encoded, not whole-string encoded). Discriminate by the presence of
      // the per-page subfolder name.
      if (url.includes("/SiteAssets/SitePages") && !url.includes("/Page1")) {
        return new Response(
          JSON.stringify({
            Folders: [{ ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/Page1", Name: "Page1" }],
            Files: [],
          }),
          { status: 200, headers: { "Content-Type": "application/json" } },
        );
      }
      if (url.includes("/Page1")) {
        return new Response(
          JSON.stringify({
            Folders: [],
            Files: [
              {
                ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/Page1/img.png",
                Name: "img.png",
                Length: 999,
                TimeLastModified: "2026-01-01T00:00:00Z",
              },
            ],
          }),
          { status: 200, headers: { "Content-Type": "application/json" } },
        );
      }
      return new Response("no match", { status: 599 });
    };
    const out = await enumerateSitePagesAssets("/sites/foo", { fetchImpl, tokenProvider });
    assert.equal(calls, 2, "expected the walker to visit both root and the Page1 subfolder");
    assert.equal(out.length, 1);
    assert.equal(out[0].name, "img.png");
  });

  it("continues the walk when a single subfolder fails", async () => {
    const fetchImpl: typeof fetch = async (input) => {
      const url = typeof input === "string" ? input : input.toString();
      if (url.includes("/SiteAssets/SitePages") && !url.includes("/Bad") && !url.includes("/Good")) {
        return new Response(
          JSON.stringify({
            Folders: [
              { ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/Bad", Name: "Bad" },
              { ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/Good", Name: "Good" },
            ],
            Files: [],
          }),
          { status: 200, headers: { "Content-Type": "application/json" } },
        );
      }
      if (url.includes("/Bad")) {
        return new Response("oops", { status: 500 });
      }
      if (url.includes("/Good")) {
        return new Response(
          JSON.stringify({
            Folders: [],
            Files: [
              {
                ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/Good/found.png",
                Name: "found.png",
                Length: 1,
                TimeLastModified: "2026-01-01T00:00:00Z",
              },
            ],
          }),
          { status: 200, headers: { "Content-Type": "application/json" } },
        );
      }
      return new Response("?", { status: 599 });
    };
    const warnings: string[] = [];
    const out = await enumerateSitePagesAssets("/sites/foo", {
      fetchImpl,
      tokenProvider,
      warn: (m) => warnings.push(m),
    });
    assert.equal(out.length, 1);
    assert.equal(out[0].name, "found.png");
    assert.equal(warnings.length, 1, "expected one warning for the bad subfolder");
    assert.match(warnings[0], /Bad.*500/);
  });

  it("flattens deeply nested folders", async () => {
    // Synthesize a chain: root → A → A/B → leaf file. Discriminate on the
    // depth of the SitePages path in the URL.
    const fetchImpl: typeof fetch = async (input) => {
      const url = typeof input === "string" ? input : input.toString();
      let body: unknown;
      if (url.includes("/SitePages/A/B")) {
        body = {
          Folders: [],
          Files: [
            {
              ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/A/B/leaf.png",
              Name: "leaf.png",
              Length: 1,
              TimeLastModified: "2026-01-01T00:00:00Z",
            },
          ],
        };
      } else if (url.includes("/SitePages/A")) {
        body = {
          Folders: [{ ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/A/B", Name: "B" }],
          Files: [],
        };
      } else if (url.includes("/SiteAssets/SitePages")) {
        body = {
          Folders: [{ ServerRelativeUrl: "/sites/foo/SiteAssets/SitePages/A", Name: "A" }],
          Files: [],
        };
      } else {
        return new Response("?", { status: 599 });
      }
      return new Response(JSON.stringify(body), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    };

    const out = await enumerateSitePagesAssets("/sites/foo", { fetchImpl, tokenProvider });
    assert.equal(out.length, 1);
    assert.equal(
      out[0].serverRelativeUrl,
      "/sites/foo/SiteAssets/SitePages/A/B/leaf.png",
    );
  });
});
