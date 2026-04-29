import { getSpoToken, SPO_ORIGIN } from "./spoTokenProvider.js";

/**
 * Tenant-wide SPO site enumeration via SP REST search.
 *
 * Uses the search index to find every site collection of class
 * `STS_Site`. Search is paged with `startrow` until exhausted.
 *
 * Filtering:
 *   - OneDrive personal (`-my.sharepoint.com`) is dropped
 *   - The app catalog and search center are dropped (operational
 *     surface, no end-user content)
 *   - Anything else is included by default — modern Communication and
 *     Team sites, classic team sites, hub sites, etc.
 *
 * Requires the calling identity to have `Sites.Read.All` (application)
 * on the SharePoint resource. The function's MI-federated app does.
 */

const TENANT_HOST = new URL(SPO_ORIGIN).hostname; // e.g. contoso.sharepoint.com

interface SearchCell {
  Key: string;
  Value: string;
}
interface SearchRow {
  Cells: SearchCell[];
}
interface SearchResult {
  PrimaryQueryResult?: {
    RelevantResults?: {
      RowCount: number;
      TotalRows: number;
      Table: {
        Rows: SearchRow[];
      };
    };
  };
}

function rowToMap(row: SearchRow): Record<string, string> {
  const out: Record<string, string> = {};
  for (const cell of row.Cells) out[cell.Key] = cell.Value;
  return out;
}

export interface SiteSummary {
  /** Server-relative path like `/sites/hub`, or `/` for the tenant root */
  serverRelativeUrl: string;
  /** Absolute URL */
  absoluteUrl: string;
  /** Site title from search */
  title: string;
  /** Web template id (e.g. `GROUP#0`, `SITEPAGEPUBLISHING#0`, `STS#0`) */
  webTemplate?: string;
  /**
   * GUID of the hub site this site is associated with, if any.
   * Returned by SP search as `DepartmentId` and is empty/zero-guid for
   * standalone sites. The site picker UI uses this to group sites
   * under their hubs.
   */
  hubSiteId?: string;
  /** True if this site is itself a hub root. */
  isHubSite: boolean;
}

/**
 * Compact hub site descriptor — used by the site picker UI to label
 * the hub group headers and offer "select all sites in this hub".
 */
export interface HubSiteSummary {
  /** Hub site identifier (GUID), matches `SiteSummary.hubSiteId`. */
  id: string;
  /** Hub site display name. */
  title: string;
  /** Server-relative path of the hub root site. */
  serverRelativeUrl: string;
}

/**
 * Heuristic filter — drop sites that aren't user content.
 *
 * Excluded:
 *   - OneDrive personal (`<tenant>-my.sharepoint.com`)
 *   - App catalog (`/sites/appcatalog`)
 *   - Search center (`/sites/search`)
 *   - Public CDN sites (`/sites/PublicCDN_*`)
 *   - The `_api`, `_layouts`, `_vti_bin` system "sites" search sometimes
 *     returns
 *
 * The filter is conservative: when in doubt, include the site. The
 * scanner is read-only and dropping a site we should have scanned is
 * worse than scanning a system site (which will just have zero pages).
 */
function isInteresting(siteUrl: string): boolean {
  const u = new URL(siteUrl);
  const host = u.hostname.toLowerCase();
  const path = u.pathname.toLowerCase();

  // OneDrive personal
  if (host.endsWith("-my.sharepoint.com")) return false;

  // System endpoints that occasionally surface as "sites"
  if (path.startsWith("/_")) return false;

  // App catalog and search center
  if (path === "/sites/appcatalog") return false;
  if (path.startsWith("/sites/appcatalog/")) return false;
  if (path === "/sites/search") return false;

  return true;
}

/**
 * Enumerate every modern site in the tenant via SP REST search.
 *
 * Returns server-relative paths suitable for the scan orchestrator
 * (e.g. `/sites/hub`, `/`).
 */
export async function enumerateSites(): Promise<SiteSummary[]> {
  const token = await getSpoToken();
  const baseUrl =
    `${SPO_ORIGIN}/_api/search/query` +
    `?querytext='contentclass:STS_Site'` +
    `&trimduplicates=false` +
    `&rowlimit=500` +
    // DepartmentId is the hub site association GUID; SiteId is the
    // site collection's own GUID (used to flag hub roots).
    `&selectproperties='Path,Title,WebTemplate,SiteId,DepartmentId'`;

  const out: SiteSummary[] = [];
  let startRow = 0;
  // Hard cap on how many search calls we'll make — defensive against
  // an infinite loop if SP misbehaves.
  let calls = 0;
  const maxCalls = 50;

  while (calls < maxCalls) {
    calls += 1;
    const url = `${baseUrl}&startrow=${startRow}`;
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
      },
    });
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`SP search failed: ${res.status} ${body.slice(0, 300)}`);
    }
    const data = (await res.json()) as SearchResult;
    const rels = data.PrimaryQueryResult?.RelevantResults;
    if (!rels) break;
    const rows = rels.Table?.Rows ?? [];
    if (rows.length === 0) break;

    for (const row of rows) {
      const m = rowToMap(row);
      const path = m.Path;
      if (!path) continue;
      try {
        const u = new URL(path);
        if (u.hostname.toLowerCase() !== TENANT_HOST) continue;
        if (!isInteresting(path)) continue;
        // Hub site association: search returns the hub's SiteId in the
        // DepartmentId field for sites that belong to a hub. A site is
        // a hub root iff its own SiteId equals its DepartmentId.
        const dept = (m.DepartmentId ?? "").replace(/^\{|\}$/g, "").toLowerCase();
        const siteId = (m.SiteId ?? "").replace(/^\{|\}$/g, "").toLowerCase();
        const empty = !dept || dept === "00000000-0000-0000-0000-000000000000";
        const hubSiteId = empty ? undefined : dept;
        const isHubSite = !empty && siteId === dept;
        out.push({
          serverRelativeUrl: u.pathname.replace(/\/$/, "") || "/",
          absoluteUrl: path,
          title: m.Title ?? "",
          webTemplate: m.WebTemplate,
          hubSiteId,
          isHubSite,
        });
      } catch {
        // skip malformed
      }
    }

    startRow += rows.length;
    if (startRow >= rels.TotalRows) break;
  }

  // De-duplicate by serverRelativeUrl (search occasionally returns
  // the same site twice with different casings).
  const seen = new Set<string>();
  const unique: SiteSummary[] = [];
  for (const s of out) {
    const k = s.serverRelativeUrl.toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    unique.push(s);
  }
  // Stable sort by URL for predictable output.
  unique.sort((a, b) => a.serverRelativeUrl.localeCompare(b.serverRelativeUrl));
  return unique;
}

/**
 * Build the list of hub-site descriptors from an already-enumerated
 * site list. We don't need a separate REST call — every hub root
 * appears in the search enumeration with `isHubSite: true`, and that's
 * all the picker UI needs (id + title + path).
 */
export function deriveHubs(sites: SiteSummary[]): HubSiteSummary[] {
  const hubs: HubSiteSummary[] = [];
  for (const s of sites) {
    if (s.isHubSite && s.hubSiteId) {
      hubs.push({
        id: s.hubSiteId,
        title: s.title || s.serverRelativeUrl,
        serverRelativeUrl: s.serverRelativeUrl,
      });
    }
  }
  hubs.sort((a, b) => a.title.localeCompare(b.title));
  return hubs;
}
