# Orphan asset cleanup: integration plan

> Sibling feature to link inventory and duplicate detection. Detects files under `SiteAssets/SitePages/` that no current page references and recycles them under the calling user's identity. Approved 2026-05-05; preserved as the design record for the implementation that followed.

## Summary

Orphan asset detection is a **derived view of the existing persistent backlinks index**, not a new scan. The link inventory scanner already extracts every URL referenced by every page on every reachable site and stores them in `link-inventory-results/backlinks-index.json`. To find orphans we just walk `SiteAssets/SitePages/` per site, compute each file's canonical key, and check whether the index already has it.

The web part gains a fourth pivot tab (**Orphans**) that matches the visual and interaction conventions of `DuplicatesTab`. The action flow mirrors `linkInventoryReplace`: admin gate, per-site OBO write check, sequential per-site execution, dry-run by default. Files move to the SharePoint recycle bin (93-day window) under the calling user's identity, so version history and recoverability are correct.

The PowerShell module from Phase 1–2 is the reference implementation for the parsing/diff logic; we are not porting it line-for-line. The TypeScript implementation reuses the existing `urlNormalizer`, `canvasLinkExtractor`, and `backlinksIndexStore` services, which already do all the parsing work.

## Why this isn't a separate scanner

Reading [func/src/services/backlinksIndex.ts](../func/src/services/backlinksIndex.ts) and [backlinksIndexStore.ts](../func/src/services/backlinksIndexStore.ts) makes the answer obvious. The persistent index has shape:

```ts
Map<canonicalKey (path-only), BacklinkEntry[]>
```

Every page in the tenant contributes its outbound links, normalized via the same `normalizeUrl` we'd use to canonicalize a `SiteAssets/SitePages/.../foo.png` server-relative path. Asking "what's orphaned?" is asking "for each file under `SiteAssets/SitePages/<site>/`, is its canonical key absent from this map?": a lookup, not a scan.

This eliminates a class of integration concerns that an independent scanner would create:
- No second site-walk loop, no duplicate queue infrastructure, no separate aggregate blob.
- The index already handles renames correctly because pages reference the actual folder URL inside `CanvasContent1`, which the parser captures faithfully. Same mechanism as the standalone PowerShell module.
- The index already handles custom-template GUID folders the same way: pages reference their own folder regardless of what it's named.
- Cross-folder references (page A in folder X references a file in folder Y) are already in the index because the parser doesn't care which folder the URL points at.

## One required upstream extension to the existing scanner

`buildPageInventory` in [linkInventoryScanner.ts](../func/src/services/linkInventoryScanner.ts) currently parses two fields:

```ts
links = [
  ...extractLinksFromCanvas(canvas),
  ...extractLinksFromCanvas(layout),
];
```

Banner images live in a third field, `BannerImageUrl` on the SitePages list item, that the canvas extractor never sees. A page's banner is commonly stored at `SiteAssets/SitePages/<page>/banner.jpg`; without scanning `BannerImageUrl`, every banner image gets false-flagged as orphan when its only referrer is the banner field.

The fix is small: extend `fetchSitePages` to also `$select` `BannerImageUrl`, and have `buildPageInventory` synthesize one extra `ExtractedLink` per page when the banner URL is non-empty. The link gets `source: "banner"` (new value in the `LinkSource` union) and flows through `classifyAll`, `canonicalKey`, and the backlinks index pipeline like any other URL. Zero impact on the rest of the system.

This is a worthwhile extension for the link inventory feature on its own merits (it surfaces broken/cross-site banner references that the current scanner misses), and it's a hard prerequisite for orphan detection to be reliable.

## Data model

No new persistent storage. Orphan detection is a query against the existing in-memory inverted index plus a freshly-walked file list per site.

```ts
interface OrphanFile {
  sitePath: string;            // /sites/foo
  serverRelativeUrl: string;   // /sites/foo/SiteAssets/SitePages/X/img.png
  fileName: string;
  pageFolder: string;          // X
  size: number;
  modified: string;            // ISO
  canonicalKey: string;        // for lookup in the existing index
}

interface OrphansReportSite {
  sitePath: string;
  siteUrl: string;
  filesScanned: number;
  orphans: OrphanFile[];
  indexAgeMs?: number;         // age of the persistent index used
  error?: string;
}

interface OrphansReport {
  generatedAt: string;
  sites: OrphansReportSite[];
  totals: { sitesScanned: number; filesScanned: number; orphans: number; };
}
```

The report is computed on demand and not persisted by default. An optional CSV export endpoint (mirroring `downloadBacklinksCsv`) exists for the audit trail; the audit row written when a recycle action runs is the durable record.

## New file enumerator: `SiteAssets/SitePages` walker

`spoFilesEnumerator.ts` already exists but explicitly skips `Site Pages`, `Style Library`, and filters `Hidden=false`, which excludes Site Assets on classic sites and on Communication sites where it hasn't been auto-provisioned. We need a separate, narrowly-scoped walker.

New service: `func/src/services/spoSitePagesAssetsEnumerator.ts`

```ts
interface SitePagesAssetFile {
  sitePath: string;
  serverRelativeUrl: string;
  fileName: string;
  size: number;
  modified: string;
}

export async function enumerateSitePagesAssets(
  sitePath: string,
): Promise<SitePagesAssetFile[]>;
```

Implementation: SP REST `getfolderbyserverrelativeurl('/sites/X/SiteAssets/SitePages')` with `$expand=Folders,Files` and recursive descent. Returns an empty array when the folder doesn't exist (sites that have never produced page assets). Failures are non-fatal at the per-site level; one bad site doesn't block the report.

The walker uses **app-only auth** (`getSpoToken`), same as the existing `spoFilesEnumerator`. Read-only enumeration doesn't need OBO; we'll filter the response by user-readable sites at the caller layer (same pattern as `filterReadableSites`).

## Endpoints

Two new HTTP triggers, mirroring the link-inventory endpoint shape:

### GET `/api/orphan-assets/report`

Generate the orphan report on demand. Loads the persistent backlinks index, per requested site walks `SiteAssets/SitePages/`, returns the report.

- Auth: `parseUserPrincipal` → `requireAdmin`
- Query/body: `sites?: string[]` (default: every site in the index), `maxConcurrency?: number` (default 6)
- Filtering: `filterReadableSites` to drop sites the calling user can't see
- Returns `OrphansReport`
- Response time on tenant-wide is a function of site count × walk cost; expect 30s–2min for hundreds of sites. Use the existing async-job pattern if it grows beyond a single response budget. For v1, sync response with concurrency cap is fine.

### POST `/api/orphan-assets/recycle`

Recycle a list of orphan files. Direct analog to `linkInventoryReplace`.

- Auth: `parseUserPrincipal` → `requireAdmin`
- Body:
  ```ts
  {
    files: Array<{ sitePath: string; serverRelativeUrl: string }>;
    dryRun?: boolean;        // default true
    confirmReportGeneratedAt?: string;  // optional staleness guard
  }
  ```
- Per-site: `getSitePermissions(user, sitePath, ctx, 'write')`. Drop sites where `canWrite=false` and record in `droppedSites`.
- Per-file: SP REST `POST .../getfilebyserverrelativeurl(...)/recycle()` using the **user's OBO write token** (`getSpoUserWriteToken`). The recycle is performed as the user; their identity is what the recycle bin records, and SP will reject the call if the user lacks delete permission.
- Returns:
  ```ts
  {
    ok: true,
    dryRun: boolean,
    summary: { totalFiles: number; recycled: number; preview: number; failed: number; droppedSites: string[] },
    results: Array<{ sitePath; serverRelativeUrl; status: 'recycled'|'preview'|'error'; error?: string }>
  }
  ```

A second helper in `linkInventoryWriter.ts` (or a new `siteAssetsWriter.ts`) wraps the recycle REST call so the OBO token plumbing is in one place. Pattern matches `applyReplacementsToPage`: take `userToken`, build the request, call SP, return a typed result.

## Audit trail

Per-recycle audit row, written to the existing job/results blob convention. Two viable options:

1. **Append-only blob** at `link-inventory-results/orphan-recycle-audit.json` with merge semantics like `backlinksIndexStore`. One row per recycled file: `{recycledAt, recycledBy, sitePath, serverRelativeUrl, fileName, dryRun, runId}`.
2. **Reuse the existing job table** by creating a synthetic "orphan-recycle" job per call. Same machinery as scan jobs: listable in the existing job UI, deletable via `deleteJob`.

Option 2 is cleaner because it's free and the UI already knows how to render job history. The downside is a low-cardinality job_kind enum gets a new value. Lean: option 2.

## Web part: new Orphans tab

Add a fourth `PivotItem` to [LinkInventoryAdmin.tsx](../spfx/link-inventory-admin/src/webparts/linkInventoryAdmin/components/LinkInventoryAdmin.tsx):

```tsx
<PivotItem headerText="Orphans" itemKey="orphans" itemIcon="DependencyRemove">
  <OrphansTab service={this.linkInventoryService} onError={this._onError} />
</PivotItem>
```

`OrphansTab` follows the established conventions from `DuplicatesTab`:

- **State shape (flat):** `loading`, `report?: IOrphansReport`, `recycling`, `expandedRows: Set<string>`, `selectedFiles: Set<string>` (for bulk action), `recentRunSummary?: ...`
- **Layout:** filter bar (site filter, "show only large files" toggle) → grouped list by `sitePath` → per-row: file name, page folder, size, modified, "View in SP" link, checkbox for selection → action bar at top with "Recycle selected (N)" disabled until selection > 0
- **Drill-in:** inline expansion under each row showing the full server-relative URL (copy button), folder context ("0 pages reference this folder"), and a "show me which page used to live here" hint when the folder name shape suggests a deleted-page provenance
- **Confirmation:** disabled-button discipline matches `DuplicatesTab`. No modal. Bulk recycle button is `disabled={recycling || selectedFiles.size === 0}`. The act of clicking is the confirmation; the result panel is the receipt.
- **Dry-run flow:** "Preview" button generates a preview report from selected files via `dryRun: true`. If the user clicks "Recycle" instead, that's a real-run direct (consistent with replace flow's separation).
- **Polling:** the report endpoint is sync for v1, no polling needed. If it grows to async, mirror the bootstrap polling pattern (2s interval, terminal-state exit).

### Service additions

Add to [LinkInventoryService.ts](../spfx/link-inventory-admin/src/webparts/linkInventoryAdmin/services/LinkInventoryService.ts):

```ts
getOrphansReport(opts?: { sites?: string[] }): Promise<IOrphansReport>;
recycleOrphans(input: {
  files: Array<{ sitePath: string; serverRelativeUrl: string }>;
  dryRun?: boolean;
  confirmReportGeneratedAt?: string;
}): Promise<IOrphansRecycleResponse>;
downloadOrphansCsv(opts?: { sites?: string[] }): Promise<void>;
```

The shape and OBO wiring follow `replace` / `getDuplicatesReport` exactly.

## Recovery surface

For v1, recovery lives in the SharePoint UI. Files go to the per-site first-stage recycle bin, then second-stage after 30 days, total 93-day retention. The Help tab gets one paragraph documenting the recovery procedure (`Restore-PnPRecycleBinItem` for power users, the SP UI for everyone else).

The audit blob/job records the exact `serverRelativeUrl` of every recycled file, so reconstructing a recovery is trivially lookup-able. An in-app "recently recycled" panel is a v2 nicety; the SP UI does the job for v1 and we don't need to duplicate it.

## Permissions

Identical to `linkInventoryReplace`:

1. Caller must be in `LINK_INVENTORY_ADMIN_GROUP_ID`
2. Per-site `getSitePermissions(user, sitePath, ctx, 'write')` for any site the user wants to recycle from. Non-write sites are silently dropped from the plan and reported in `droppedSites`.
3. Recycle calls use the user's OBO write token, so SP enforces the user's actual permissions on the file. App-only is read-only here.

The Entra app needs `AllSites.Write` (Delegated) which it already has; that's what the replace path uses.

## Edge cases

| Case | Handling | Source |
|---|---|---|
| Renamed pages | Already handled by reference scanning. The page's `CanvasContent1` keeps the original folder URL after rename. | PowerShell validation, MS Q&A |
| Custom-template GUID folders | Already handled. `CanvasContent1` references whatever folder the page uses, GUID or otherwise. | PowerShell validation |
| `BannerImageUrl` references | **Requires the upstream parser extension** described above. Without it, banner images false-flag as orphan. | Discovered while reading the existing scanner |
| Translation pages | Each translation is a separate Site Pages list item with its own `CanvasContent1`. Index sees them all; a shared asset stays referenced as long as any translation references it. | No new code |
| Drafts / checked-out pages | The scanner reads the published `CanvasContent1`. Pages mid-edit with new asset references won't have those references in the index until published. **Documented limitation;** recycle window covers this risk (93 days). | Existing behavior, document it |
| Custom SPFx web parts with non-string asset refs | Anything not URL-shaped in the JSON walker won't be captured. **Documented limitation** in v1; can add web-part-specific extractors as needed. | Existing limitation |
| Site Pages `Templates/` subfolder | Confirmed already handled. `fetchSitePages` queries `getbytitle('Site Pages')/items` with no folder filter, so list items in the `Templates/` subfolder are returned alongside top-level pages and contribute their references to the index. | [spoPagesClient.ts:201](../func/src/services/spoPagesClient.ts#L201) |
| Site Assets library missing | Some sites (esp. fresh Communication sites) have no Site Assets library yet. Walker returns empty array; report shows `filesScanned: 0` for those sites. No-op. | PowerShell validation |
| User has scan-time access but lost write access at recycle time | OBO permission check at recycle time uses live SP perms, not scan-time perms. Site is dropped from the plan. | Existing replace pattern |

## Test plan

Function-side, mirroring the existing `func/src/__tests__/` pattern (Pester equivalent: Vitest/Node test with the `--import` setup file):

- `spoSitePagesAssetsEnumerator.test.ts`: fixture `_api/web/getfolderbyserverrelativeurl` responses, recursion, missing-folder graceful handling.
- `orphanQuery.test.ts`: pure function: given a fake in-memory `BacklinksIndex` and a fake file list, returns the expected orphan diff. Exhaustive for renamed-page, custom-template, cross-folder, BannerImageUrl, and stray-upload scenarios.
- `orphansEndpoint.test.ts`: handler-level test with mocked auth + mocked enumerator + mocked index store.
- `orphanRecycleEndpoint.test.ts`: handler-level: admin gate, per-site write check, dry-run vs real, partial failures, dropped-sites reporting.

The Phase 1–2 PowerShell test fixture transfers conceptually to Vitest fixtures: same five scenarios (normal page, renamed page, deleted page asset folder, manual stray, cross-folder reference) become test inputs to `orphanQuery`. We've validated the expected outputs against a real tenant; that's our oracle.

End-to-end against a real site can re-use the existing dev tenant scratchpad (`https://lgwxt.sharepoint.com/sites/sporphan-test`); the fixture is still live there.

## Out of scope for v1

- **Other `SiteAssets/` subtrees** (logos, theme assets, list attachments). Different lifecycle, different mental model. Could be a v2 separate report; not in this feature.
- **Empty folders left behind after recycling files.** SharePoint doesn't auto-clean empty folders; we don't either. Re-running the scan returns no orphans (idempotent), and the empty folders are visually obvious in SP if a content owner wants to clear them.
- **In-app recovery UI.** SP UI is sufficient; the audit job records what to look for.
- **Scheduled / unattended runs.** v1 is admin-driven only. A timer trigger that auto-recycles is too dangerous without a cooling-off window; if requested, structure as "preview → email → human approves → execute" rather than fully automated.

## Resolved decisions

1. **`BannerImageUrl` extension: bundled with this feature.** First step of the implementation order. Touches `fetchSitePages`, `buildPageInventory`, the `LinkSource` union, and adds a unit test for banner URLs flowing through `classifyAll`. The link inventory tab benefits immediately (broken/cross-site banners surface in the existing report); orphan detection becomes reliable.
2. **Scan freshness gate: short window with override.** The README's design has a daily delta scan, so the index normally trails real edits by < 24h. A page published this morning whose banner asset gets recycled this afternoon is a bad day even with the 93-day recycle window. Proposed defaults:
    - Warn (yellow message bar) at **> 1h** since last index merge
    - Block (refuse with 409) at **> 6h** since last index merge, unless the request body sets `acknowledgedStaleIndex: true`
    - Both thresholds configurable via env vars (`ORPHAN_INDEX_WARN_AGE_MIN=60`, `ORPHAN_INDEX_BLOCK_AGE_MIN=360`)
    - Computed from the persistent index's top-level `latestAt` (max of all `sourceUpdatedAt` values), already exposed by `inflateToMemory`
    - Tighten further if the dev tenant tells us 1h is too generous in practice; I'd rather start tight and loosen than the reverse.
3. **Audit storage: synthetic orphan-recycle jobs in the existing job table.** Reuses the existing list/show/delete UI, gets the 30-day retention purge for free (which doubles as automatic audit-log expiry, fine for typical compliance windows), and adds no new persistent surface. Trade-off acknowledged: long-term forensic recovery beyond 30 days requires either changing job retention globally or shipping an append-only audit blob in v1.5. Not worth blocking v1 on.
4. **Site Pages `Templates/` subfolder: already handled.** `fetchSitePages` queries the list items endpoint with no folder filter; templates are returned and their `CanvasContent1` references flow into the index alongside top-level pages. No code change needed.
5. **Standalone PowerShell module: keep it published.** Single-site users (Reddit-thread audience) get a runnable artifact; tenant-scale users get the integrated version. The PowerShell module's parsing logic doesn't need to track changes to the TypeScript scanner because it's a one-shot tool, so divergence cost is bounded. Phase 3.5 ships it as a small intentionally-minimal GitHub repo (MIT, README, the existing module + tests, no CI). Blog post links to both.
6. **Recycle helper: new file.** `func/src/services/siteAssetsWriter.ts`. Page-edit-shaped vs file-shaped responsibilities stay separate; type signatures don't pollute `linkInventoryWriter.ts`.

## Implementation order (proposed, after plan approval)

1. **Upstream parser extension:** add `BannerImageUrl` to `fetchSitePages` $select + extend `buildPageInventory` to synthesize a banner link. Add `LinkSource: 'banner'` to the union. Add unit test covering banner URL flowing through `classifyAll` → canonical key. (Existing scanner gets this benefit immediately.)
2. **`spoSitePagesAssetsEnumerator`** + tests.
3. **Orphan query** as a pure function over `(BacklinksIndex, SitePagesAssetFile[])`. Tests transferred from PS fixture.
4. **`/api/orphan-assets/report`** endpoint + handler test.
5. **Recycle helper** in new `siteAssetsWriter.ts` + handler-level test.
6. **`/api/orphan-assets/recycle`** endpoint + handler test.
7. **Service additions** in `LinkInventoryService.ts` (sync, no polling).
8. **`OrphansTab`** + Pivot wiring.
9. **Help tab** copy update.
10. **End-to-end test** on `sporphan-test` dev tenant.
