# Link Inventory Admin

Tenant-wide SharePoint link inventory, duplicate-file detection, and orphan-asset management web part.

## Overview

Link Inventory Admin is an SPFx web part that provides three complementary reports for tenant governance:

1. **Link Inventory** — Extract and search all hyperlinks from modern pages and Office/PDF documents across the tenant. Filter by link class (OnPremises, sharing links, external URLs, etc.), find broken links, and perform tenant-wide find-and-replace operations.

2. **Duplicates** — Identify exact-match file duplicates via SHA-256 hashing, detect stale copies by comparing file versions, and flag same-name files across sites for manual triage.

3. **Orphans** — Discover unreferenced files under `SiteAssets/SitePages/` (typically left behind when pages are deleted) and recycle them to the per-site bin with full audit attribution.

## How It Works

The web part is a React+TypeScript SPFx solution (v1.22.0) that communicates with the `link-inventory-func` Azure Function via delegated Entra ID (On-Behalf-Of token flow). All scans run as background jobs against Azure Storage — pages and documents are queued per-site, processed asynchronously, and results are persisted to Blob Storage. The UI polls job status and renders aggregated findings once complete.

**Key components:**

- **LinkInventoryAdminWebPart.ts** — Main web part class; manages property pane configuration (Azure Function URL and Entra app api:// URI).
- **LinkInventoryAdmin.tsx** — Root component; renders the pivot-tab UI and error bars.
- **LinkInventoryTab.tsx** — Search and filter inventory by site, link class, source web part, and free text; trigger new scans; manage daily schedule; perform find-and-replace and backlinks operations.
- **DuplicatesTab.tsx** — Query the duplicates report; manage an allowlist to suppress expected duplicates; bootstrap the index from file version history.
- **OrphansTab.tsx** — Generate orphan reports and batch-recycle files to the recycle bin; respects calling user's write permissions via OBO.
- **LinkInventoryService.ts** — TypeScript service layer wrapping all Azure Function REST endpoints; handles AadHttpClient setup and URL construction.

The web part uses **Fluent UI (Office Fabric)** for all UI components and runs against **SPFx 1.22.0** with React 17.

## Configuration

**Web Part Properties** (set via property pane after deployment):

- **functionUrl** — Base URL of the deployed Azure Function (e.g., `https://func-linkinv01.azurewebsites.net`), no trailing slash.
- **entraAppApiUri** — Entra app's Application ID URI in `api://` form (e.g., `api://01a2adc7-7283-4208-bfe3-34accc1bf43a`). Found on the app registration's "Expose an API" page. Do not use the friendly display name; use the api:// form to avoid intermittent resolution failures.

**Preconfigured Entries** (in manifest):

- Group ID: `cf066440-3577-4f34-aea2-2e03f6d4cda4` (Tools group)
- Default title: "Link Inventory Admin"
- Default description: "Browse the tenant-wide link inventory and duplicate-file report"

## Dependencies

**Graph Permissions** (from `package-solution.json`):

- `link-inventory-func` / `user_impersonation` — Delegated permission for OBO token flow to the Azure Function.

**Related Solutions in Portfolio:**

This solution is the admin interface for the Link Inventory system. Separate deployments:

- **spfx-backlinks-column** — Field customizer that renders the Backlinks column created by this web part's `enableBacklinksColumn` API call. Deploy separately and update the function configuration.
- **redirect-manager** suite — Shares the same Entra app registration (`altra-spo-link-inventory`, api://01a2adc7-7283-4208-bfe3-34accc1bf43a) and the same Azure Function app.

**SharePoint Lists/Libraries:**

The web part does not directly read or write SharePoint lists; all interactions go through the Azure Function. The backend manages:

- A persistent backlinks index table in Azure Storage.
- Scan job metadata and results blobs (keyed by jobId).
- Duplicates bootstrap jobs and allowlist entries.

## Deployment

**Build:**

```bash
npm install
npm run build
```

Produces `./release/link-inventory-admin.sppkg` (SPFx solution package).

**Deploy to Tenant App Catalog:**

1. Upload `link-inventory-admin.sppkg` to the tenant-wide app catalog.
2. Enable "Make this solution available to all sites in your organization."
3. Add the web part to a SharePoint page (typically in a designated admin site).
4. Open the web part's property pane and configure:
   - Azure Function URL (e.g., `https://func-linkinv01.azurewebsites.net`)
   - Entra app api:// URI (confirm with your Azure AD admin)

**Site Deployment:**

```bash
npm run build:dev  # Local dev build
npm run start      # Serve at https://localhost:4321
```

## API Endpoints (via LinkInventoryService)

All endpoints require OBO authentication to the configured Entra app. Common patterns:

- `GET /api/link-inventory/whoami` — Current user info and admin check.
- `GET /api/link-inventory/sites` — Tenant site list (cached 5 min).
- `POST /api/link-inventory/scan` — Trigger a page scan.
- `POST /api/link-inventory/scan-docs` — Trigger a document scan (with preview option).
- `POST /api/link-inventory/scan-unified` — Trigger paired page + document scan.
- `GET /api/link-inventory/scan/<jobId>` — Fetch job status.
- `GET /api/link-inventory/scan/<jobId>/results` — Fetch scan results blob.
- `POST /api/link-inventory/replace` — Perform find-and-replace on a job's results.
- `GET /api/link-inventory/schedule` / `PUT /api/link-inventory/schedule` — Configure daily delta scans.
- `GET /api/duplicates/report` — Fetch duplicates report.
- `GET /api/duplicates/allowlist` — Fetch allowlist.
- `POST /api/duplicates/allowlist` — Add allowlist entry (hash, path pattern, or name pattern).
- `GET /api/orphan-assets/report` — Generate orphans report.
- `POST /api/orphan-assets/recycle` — Batch-recycle orphan files.
- `GET /api/link-inventory/backlinks/batch` — Batch backlinks lookup (chunked, up to 200 fileRefs per call).

## Known Issues / Caveats

1. **Entra app api:// URI resolution** — Always pass the `api://` form of the Application ID URI rather than the friendly display name. The friendly-name lookup intermittently fails because it goes through SharePoint's approved-permissions list and can time out or resolve incorrectly.

2. **Stale index warnings on orphan reports** — The orphans report uses a persistent backlinks index built from the most recent completed scans. If the index is stale (older than a configurable threshold, default 60 min for warnings, 360 min for hard blocks), the UI will display a yellow warning or block the report until the user acknowledges the risk. This is because pages added since the last scan will false-flag as orphans.

3. **Per-site read filtering** — Results endpoints return only sites the calling user can read. Sites the user lacks access to appear in `droppedSites` so the UI can show "8 of 12 sites visible to you." Non-admin users cannot trigger scans; the UI hides the button via a whoami probe, but the backend enforces 403 if accessed directly.

4. **Document scan file size limits** — Document scans apply a configurable max file size (default 100 MB, hard ceiling 500 MB). Files larger than the limit are skipped during enumeration.

5. **Preview-only doc scans** — The document scan endpoint supports `previewOnly: true` to enumerate files and return per-site counts without downloading content. These preview jobs can be promoted to full scans later via `promoteDocScan`.

6. **Backlinks column requires separate deployer** — The `enableBacklinksColumn` API creates a stub `RmgrBacklinks` column and registers a field customizer, but the customizer itself is a separate SPFx package (`spfx-backlinks-column`) and must be deployed independently. Until deployed, the column renders as plain text.

7. **Duplicates allowlist filtering** — The allowlist supports three entry types: exact SHA-256 hashes (byte-for-byte duplicates), path regex patterns, and name patterns. Entries suppress rows from the duplicates report but do not retroactively modify the underlying scanned data.

8. **SharePoint navigation interception** — DuplicatesTab and OrphansTab render external links with `data-interception="off"` to prevent SharePoint's modern page router from intercepting same-origin clicks. Without this, links would open in-place instead of in a new tab.

## Solution Metadata

| Field | Value |
|-------|-------|
| Name | Link Inventory Admin |
| Solution ID | d97f31aa-9484-45a1-bc14-575baf3887df |
| Version | 0.1.4.0 (package.json: 0.1.3) |
| Feature ID | 2561f528-a261-453a-a339-776be2df55b8 |
| Web Part ID | 28568c17-784e-43f1-b91b-a8e21f345733 |
| SPFx Framework | 1.22.0 |
| React | 17.0.1 |
| Node.js | >=18.17.1 <23.0.0 |
| TypeScript | ~5.8.0 |
| Build Tool | Heft + SPFx build rig |
| Package Scope | `spfx-link-inventory-admin` (private) |

**Source Repository:** https://github.com/charlie-vogt/spo-link-inventory

**Tenant Scope:** Organization-wide app catalog required; web part can be added to any site after approval.
