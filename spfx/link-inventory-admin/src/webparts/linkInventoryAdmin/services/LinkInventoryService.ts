import { AadHttpClient, type AadHttpClientFactory } from '@microsoft/sp-http';

/**
 * Service wrapper for the link-inventory function endpoints.
 *
 * All endpoints require a user bearer token (via OBO), so they go
 * through `AadHttpClient`. The constructor takes the Entra app's
 * api:// URI as a configuration parameter — that URI is what
 * `aadHttpClientFactory.getClient(...)` uses to obtain a delegated
 * token against the right resource. Pass it as the `api://<guid>`
 * form rather than the friendly display name; the friendly-name
 * lookup goes through SP's approved-permissions list and
 * intermittently fails to resolve.
 */

export interface ILinkInventoryWhoami {
  ok: boolean;
  userId: string;
  upn?: string;
  isAdmin: boolean;
}

export interface ISiteSummary {
  serverRelativeUrl: string;
  absoluteUrl: string;
  title: string;
  webTemplate?: string;
  hubSiteId?: string;
  isHubSite: boolean;
}

export interface IHubSummary {
  id: string;
  title: string;
  serverRelativeUrl: string;
}

/**
 * Daily-scan schedule config — round-trip shape for
 * GET/PUT /api/link-inventory/schedule.
 */
export interface IScheduleConfig {
  enabled: boolean;
  /** "HH:MM" 24-hour in `timeZone`. */
  timeOfDay: string;
  /** IANA tz name (default America/Chicago). */
  timeZone: string;
  /** YYYY-MM-DD (in `timeZone`) of the last successful fire. Empty before first. */
  lastFiredOnDate: string;
  /** Page-scan job id from the last successful fire. */
  lastJobId: string;
  /** ISO timestamp of the last successful fire. */
  lastFiredAt: string;
  /** ISO timestamp of the last config edit. */
  updatedAt: string;
  /** UPN of the last admin who edited. */
  updatedBy: string;
}

/** PUT body — partial update; missing fields stay as-is. */
export interface IScheduleConfigInput {
  enabled?: boolean;
  timeOfDay?: string;
  timeZone?: string;
}

/**
 * Per-extension breakdown inside an included/excluded bucket.
 */
export interface IPreviewBucketStats {
  files: number;
  totalBytes: number;
  byExtension: Record<string, { count: number; totalBytes: number }>;
}

/**
 * Per-site/per-library file counts returned by a preview-only doc scan.
 */
export interface IPreviewSiteSummary {
  site: string;
  libraries: Record<string, { files: number; totalBytes: number }>;
  totalFiles: number;
  totalBytes: number;
  /** Present on scans from the duplicate-detection era. */
  included?: IPreviewBucketStats;
  /** Present on scans from the duplicate-detection era. */
  excluded?: IPreviewBucketStats;
}

export interface IPreviewResults {
  jobId: string;
  kind: 'documents-preview';
  startedAt: string;
  finishedAt: string;
  previewOnly: true;
  siteSummaries: IPreviewSiteSummary[];
  totals: { sites: number; files: number; bytes: number };
  /** Tenant-wide bucket totals — present on newer preview blobs. */
  included?: IPreviewBucketStats;
  /** Tenant-wide bucket totals — present on newer preview blobs. */
  excluded?: IPreviewBucketStats;
}

export interface ILinkInventoryJob {
  jobId: string;
  /** "pages" (modern Site Pages) or "documents" (OOXML/PDF). */
  kind?: 'pages' | 'documents';
  status: 'queued' | 'running' | 'completed' | 'failed';
  startedAt: string;
  finishedAt?: string;
  sites: string[];
  sitesTotal: number;
  sitesCompleted: number;
  currentSite?: string;
  pagesTotal: number;
  linksTotal: number;
  /** Document scans only. */
  filesTotal?: number;
  /** Document scans only. */
  filesCompleted?: number;
  /** Document scans only — true if this was a preview-only enumeration job. */
  previewOnly?: boolean;
  errorCount: number;
  recentErrors: string[];
  resultsAvailable: boolean;
  caller?: string;
  /** When set, this job is one half of a unified scan paired with the
   *  job at this id. Used by the UI to present them as a single entry. */
  siblingJobId?: string;
}

export interface IClassifiedLink {
  rawUrl: string;
  url: string;
  source: string;
  webPartInstanceId?: string;
  text?: string;
  linkClass: string;
  normalizedKey: string;
  /**
   * Set only when `linkClass === "malformed-spo-link"`. Tells the UI
   * why the link is flagged: `search-fragment` (leftover `#search=`
   * from a search-UI click), `file-not-found` (opt-in verifier saw
   * 404 on the `id=` path), or `both`.
   */
  malformedReason?: 'search-fragment' | 'file-not-found' | 'both';
  /** Suggested replacement URL — populated for sharing links and Office Online wrappers. */
  suggestion?: string;
  /**
   * Canonical equivalence key — same string for any URL form pointing
   * at the same target. Used by the inventory UI to group "this link
   * is also at N other places" in the detail panel.
   */
  canonicalKey?: string;
}

export interface IPageInventory {
  pageId: number;
  pageTitle: string;
  pageUrl: string;
  modified?: string;
  etag?: string;
  links: IClassifiedLink[];
  parseError?: string;
  /**
   * Set client-side when the page came in via a unified-scan merge.
   * Tells the inventory which scan kind ("pages" or "documents") this
   * row originated from so the table can render a Source column and
   * users can filter by source. Server never sets this — the server's
   * `getResults` blob is per-kind and untagged.
   */
  sourceKind?: 'pages' | 'documents';
}

export interface ISiteInventory {
  site: string;
  pageCount: number;
  linkCount: number;
  byClass: Record<string, number>;
  bySource: Record<string, number>;
  pages: IPageInventory[];
  scanMs: number;
  error?: string;
}

export interface ILinkInventoryResults {
  jobId: string;
  startedAt: string;
  finishedAt: string;
  sites: ISiteInventory[];
  totals: { sites: number; pages: number; links: number; errors: number };
  droppedSites: string[];
  userScopedTotals: { sites: number; pages: number; links: number; errors: number };
}

// AadHttpClient resource: pass the Entra app's Application ID URI
// (api://<guid>) rather than the friendly display name. Friendly-name
// lookup goes through SP's approved-permissions list and intermittently
// fails to resolve to a valid scope; the URI form passes through directly.

export class LinkInventoryService {
  private aadClientPromise: Promise<AadHttpClient> | undefined;

  public constructor(
    private aadHttpClientFactory: AadHttpClientFactory,
    private functionUrl: string,
    /** Entra app api:// URI used as the OBO target resource. */
    private entraAppApiUri: string,
  ) {}

  private async _getClient(): Promise<AadHttpClient> {
    if (!this.aadClientPromise) {
      if (!this.entraAppApiUri) {
        throw new Error('Entra app api:// URI is not configured. Set it in the web part property pane.');
      }
      this.aadClientPromise = this.aadHttpClientFactory.getClient(this.entraAppApiUri);
    }
    return this.aadClientPromise;
  }

  private _url(path: string, query?: Record<string, string>): string {
    const base = this.functionUrl.replace(/\/$/, '') + path;
    const params = new URLSearchParams();
    if (query) {
      const keys = Object.keys(query);
      for (let i = 0; i < keys.length; i++) {
        params.append(keys[i], query[keys[i]]);
      }
    }
    const qs = params.toString();
    return qs ? `${base}?${qs}` : base;
  }

  private async _get<T>(path: string, query?: Record<string, string>): Promise<T> {
    const client = await this._getClient();
    const res = await client.get(this._url(path, query), AadHttpClient.configurations.v1);
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`GET ${path} failed (${res.status}): ${text}`);
    }
    return (await res.json()) as T;
  }

  private async _post<T>(path: string, body: unknown): Promise<T> {
    const client = await this._getClient();
    const res = await client.post(this._url(path), AadHttpClient.configurations.v1, {
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`POST ${path} failed (${res.status}): ${text}`);
    }
    return (await res.json()) as T;
  }

  private async _delete<T>(path: string): Promise<T> {
    const client = await this._getClient();
    // AadHttpClient exposes `fetch()` for arbitrary methods. We use it
    // here because there's no `delete()` shortcut.
    const res = await client.fetch(this._url(path), AadHttpClient.configurations.v1, {
      method: 'DELETE',
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`DELETE ${path} failed (${res.status}): ${text}`);
    }
    return (await res.json()) as T;
  }

  private async _put<T>(path: string, body: unknown): Promise<T> {
    const client = await this._getClient();
    const res = await client.fetch(this._url(path), AadHttpClient.configurations.v1, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`PUT ${path} failed (${res.status}): ${text}`);
    }
    return (await res.json()) as T;
  }

  public async whoami(): Promise<ILinkInventoryWhoami> {
    return this._get<ILinkInventoryWhoami>('/api/link-inventory/whoami');
  }

  public async getSchedule(): Promise<IScheduleConfig> {
    const res = await this._get<{ ok: true; config: IScheduleConfig }>('/api/link-inventory/schedule');
    return res.config;
  }

  public async setSchedule(input: IScheduleConfigInput): Promise<IScheduleConfig> {
    const res = await this._put<{ ok: true; config: IScheduleConfig }>('/api/link-inventory/schedule', input);
    return res.config;
  }

  public async listJobs(): Promise<ILinkInventoryJob[]> {
    const res = await this._get<{ jobs: ILinkInventoryJob[] }>('/api/link-inventory/scan');
    return res.jobs;
  }

  public async getJob(jobId: string): Promise<ILinkInventoryJob> {
    const res = await this._get<{ job: ILinkInventoryJob }>(`/api/link-inventory/scan/${jobId}`);
    return res.job;
  }

  /**
   * Delete a scan job and every artifact it produced (table row,
   * aggregate result blob, per-site partials, manifests). Admin-only.
   *
   * Returns `{ rowDeleted, blobsDeleted }` so the caller can show a
   * "Deleted N blobs" toast.
   */
  public async deleteJob(jobId: string): Promise<{ rowDeleted: boolean; blobsDeleted: number }> {
    return this._delete<{ rowDeleted: boolean; blobsDeleted: number }>(
      `/api/link-inventory/scan/${jobId}`,
    );
  }

  public async triggerScan(options?: {
    sites?: string[];
    pageIdsBySite?: Record<string, number[]>;
    /**
     * Opt-in: after scanning, HEAD-check every AllItems.aspx `?id=` path
     * against SP REST. 404s get upgraded from `spo-internal` to
     * `malformed-spo-link` (reason `file-not-found`). Adds one REST
     * call per unique AllItems link — costly on large scans.
     */
    verifyFiles?: boolean;
  }): Promise<{ jobId: string; sitesTotal: number; sites: string[] }> {
    return this._post('/api/link-inventory/scan', options ?? {});
  }

  /**
   * Trigger a document scan. Files are filtered by `maxFileBytes`
   * (default 100 MB, hard ceiling 500 MB) and optionally by
   * `modifiedAfter` for incremental rescans.
   *
   * Pass `fileRefs` to scan specific files only (skips library
   * enumeration entirely). Pass `sites` to scan specific sites only.
   * Pass neither to scan everything.
   *
   * Pass `previewOnly: true` to enumerate files and return per-site
   * counts without actually downloading or scanning anything. The
   * resulting job can be promoted to a real scan via promoteDocScan.
   */
  public async triggerDocScan(options?: {
    sites?: string[];
    fileRefs?: string[];
    maxFileBytes?: number;
    modifiedAfter?: string;
    previewOnly?: boolean;
    /** Opt-in SPO file verification — see triggerScan. */
    verifyFiles?: boolean;
  }): Promise<{ jobId?: string; kind?: string; filesTotal?: number; sitesTotal?: number; message?: string; enumErrors?: string[] }> {
    return this._post('/api/link-inventory/scan-docs', options ?? {});
  }

  /**
   * Promote a completed preview-only doc scan to a real scan,
   * optionally narrowing to a subset of sites.
   */
  public async promoteDocScan(previewJobId: string, sites?: string[]): Promise<{
    jobId: string;
    sitesTotal: number;
    filesTotal: number;
  }> {
    return this._post('/api/link-inventory/scan-docs/promote', { previewJobId, sites });
  }

  /**
   * Trigger a paired page-scan + doc-scan together. Returns both jobIds
   * so the UI can poll either side and present them as one scan.
   */
  public async triggerUnifiedScan(options?: {
    sites?: string[];
    enumerateAll?: boolean;
    maxFileBytes?: number;
    modifiedAfter?: string;
    verifyFiles?: boolean;
  }): Promise<{
    ok: boolean;
    pageJobId: string;
    docJobId: string;
    sitesTotal: number;
    enumerated?: boolean;
    invalid?: string[];
  }> {
    return this._post('/api/link-inventory/scan-unified', options ?? {});
  }

  /**
   * Fetch the tenant site list (with hub associations) for the site
   * picker UI. Cached server-side for 5 minutes.
   */
  public async listSites(): Promise<{ sites: ISiteSummary[]; hubs: IHubSummary[] }> {
    return this._get<{ sites: ISiteSummary[]; hubs: IHubSummary[] }>('/api/link-inventory/sites');
  }

  public async getResults(jobId: string): Promise<ILinkInventoryResults> {
    const res = await this._get<{ results: ILinkInventoryResults }>(`/api/link-inventory/scan/${jobId}/results`);
    return res.results;
  }

  /**
   * Fetch results for a preview-only doc scan. Same endpoint as
   * getResults, but the blob shape is different (per-site file counts
   * rather than per-page link inventory) so we return a different type
   * to keep callers honest.
   */
  public async getPreviewResults(jobId: string): Promise<IPreviewResults> {
    const res = await this._get<{ results: IPreviewResults }>(`/api/link-inventory/scan/${jobId}/results`);
    return res.results;
  }

  public async replace(
    jobId: string,
    find: string,
    replace: string,
    options?: { dryRun?: boolean; sites?: string[]; pageIds?: number[]; mode?: 'substring' | 'canonical' },
  ): Promise<ILinkInventoryReplaceResponse> {
    return this._post<ILinkInventoryReplaceResponse>('/api/link-inventory/replace', {
      jobId,
      find,
      replace,
      dryRun: options?.dryRun !== false,
      sites: options?.sites,
      pageIds: options?.pageIds,
      mode: options?.mode,
    });
  }

  /**
   * Enable the Backlinks column on a library. Backend creates the stub
   * `RmgrBacklinks` column, registers a field customizer on it, and adds
   * the column to the default view.
   *
   * The field customizer (`spfx-backlinks-column`) is a separate SPFx
   * package and is not shipped with this admin web part — see the project
   * README for the deploy guide. Until that customizer is deployed the
   * column is created but renders as a plain text field.
   *
   * Throws 409 if the target library already has a column named
   * "Backlinks" or "RmgrBacklinks" — admin must rename/remove it first.
   */
  public async enableBacklinksColumn(
    siteUrl: string,
    libraryTitle: string,
  ): Promise<IEnableBacklinksColumnResponse> {
    return this._post<IEnableBacklinksColumnResponse>('/api/link-inventory/backlinks-column/enable', {
      siteUrl,
      libraryTitle,
      customizerFunctionUrl: this.functionUrl,
      customizerFunctionKey: '',
    });
  }

  /**
   * Rebuild the persistent backlinks index from the most recent
   * completed page-scan + doc-scan. Admin-only. Used for bootstrap
   * on first deploy and for recovery if the index gets out of sync.
   */
  public async rebuildBacklinksIndex(): Promise<IRebuildBacklinksResponse> {
    return this._post<IRebuildBacklinksResponse>('/api/link-inventory/backlinks-index/rebuild', {});
  }

  /**
   * Download the backlinks CSV. Pass `siteUrl` + `libraryTitle` for a
   * per-library export (one row per file × backlink); omit both for a
   * tenant-wide dump of the inverted index. Triggers a browser
   * download via a blob URL. Admin-only on the backend.
   */
  public async downloadBacklinksCsv(opts?: { siteUrl?: string; libraryTitle?: string }): Promise<void> {
    const client = await this._getClient();
    const query: Record<string, string> = {};
    if (opts?.siteUrl) query.siteUrl = opts.siteUrl;
    if (opts?.libraryTitle) query.libraryTitle = opts.libraryTitle;
    const url = this._url('/api/link-inventory/backlinks/export', query);
    const res = await client.get(url, AadHttpClient.configurations.v1);
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Backlinks CSV export failed (${res.status}): ${text.slice(0, 300)}`);
    }
    const disposition = res.headers.get('Content-Disposition') ?? '';
    const match = /filename="([^"]+)"/.exec(disposition);
    const filename = match?.[1] ?? `backlinks-${new Date().toISOString().slice(0, 10)}.csv`;
    const blob = await res.blob();
    const href = URL.createObjectURL(blob);
    try {
      const a = document.createElement('a');
      a.href = href;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
    } finally {
      // Revoke after a tick so the browser has a chance to start the
      // download before we invalidate the object URL.
      setTimeout(() => URL.revokeObjectURL(href), 2000);
    }
  }

  public async getDuplicatesReport(caps?: {
    maxExactGroups?: number;
    maxStalePairs?: number;
    maxSameNamePairs?: number;
    maxNearDuplicatePairs?: number;
    maxDivergedPairs?: number;
  }): Promise<IDuplicatesReport> {
    const query: Record<string, string> = {};
    if (caps?.maxExactGroups) query.maxExactGroups = String(caps.maxExactGroups);
    if (caps?.maxStalePairs) query.maxStalePairs = String(caps.maxStalePairs);
    if (caps?.maxSameNamePairs) query.maxSameNamePairs = String(caps.maxSameNamePairs);
    if (caps?.maxNearDuplicatePairs) query.maxNearDuplicatePairs = String(caps.maxNearDuplicatePairs);
    if (caps?.maxDivergedPairs) query.maxDivergedPairs = String(caps.maxDivergedPairs);
    return this._get<IDuplicatesReport>('/api/duplicates/report', query);
  }

  public async getDuplicatesForFile(fileRef: string): Promise<IDuplicatesLookupResponse> {
    return this._get<IDuplicatesLookupResponse>('/api/duplicates/lookup', { fileRef });
  }

  public async getDuplicatesAllowlist(): Promise<{ ok: boolean; isAdmin: boolean; allowlist: IDuplicatesAllowlist }> {
    return this._get<{ ok: boolean; isAdmin: boolean; allowlist: IDuplicatesAllowlist }>('/api/duplicates/allowlist');
  }

  public async addDuplicatesAllowlistEntry(input: {
    kind: 'hash' | 'path' | 'name';
    sha256?: string;
    pattern?: string;
    note: string;
  }): Promise<{ ok: boolean; allowlist: IDuplicatesAllowlist }> {
    return this._post<{ ok: boolean; allowlist: IDuplicatesAllowlist }>('/api/duplicates/allowlist', input);
  }

  public async removeDuplicatesAllowlistEntry(input: {
    kind: 'hash' | 'path' | 'name';
    sha256?: string;
    pattern?: string;
  }): Promise<{ ok: boolean; allowlist: IDuplicatesAllowlist }> {
    const client = await this._getClient();
    const res = await client.fetch(this._url('/api/duplicates/allowlist'), AadHttpClient.configurations.v1, {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(input),
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`DELETE /api/duplicates/allowlist failed (${res.status}): ${text}`);
    }
    return (await res.json()) as { ok: boolean; allowlist: IDuplicatesAllowlist };
  }

  public async downloadDuplicatesCsv(type: 'exact' | 'stale' | 'samename'): Promise<void> {
    const client = await this._getClient();
    const url = this._url('/api/duplicates/export', { type });
    const res = await client.get(url, AadHttpClient.configurations.v1);
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Duplicates CSV export failed (${res.status}): ${text.slice(0, 300)}`);
    }
    const disposition = res.headers.get('Content-Disposition') ?? '';
    const match = /filename="([^"]+)"/.exec(disposition);
    const filename = match?.[1] ?? `duplicates-${type}-${new Date().toISOString().slice(0, 10)}.csv`;
    const blob = await res.blob();
    const href = URL.createObjectURL(blob);
    try {
      const a = document.createElement('a');
      a.href = href;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
    } finally {
      setTimeout(() => URL.revokeObjectURL(href), 2000);
    }
  }

  public async bootstrapDuplicatesFromVersions(input: {
    sitePath: string;
    libraryTitle: string;
    maxFiles?: number;
    maxVersionsPerFile?: number;
    includeOther?: boolean;
  }): Promise<IDuplicatesBootstrapStartResponse> {
    return this._post<IDuplicatesBootstrapStartResponse>('/api/duplicates/bootstrap', input);
  }

  /**
   * Batch backlinks lookup. The endpoint caps each call at 200 fileRefs,
   * so this method chunks larger inputs and merges the results into a
   * single record keyed by fileRef. Empty `fileRefs` returns an empty
   * map without making a request.
   */
  public async getBacklinksBatch(fileRefs: string[]): Promise<IBacklinksBatchResponse> {
    if (fileRefs.length === 0) {
      return { isAdmin: false, results: {} };
    }
    const CHUNK = 200;
    const merged: IBacklinksBatchResponse = { isAdmin: false, results: {} };
    let mergedAdmin = false;
    let mergedBuiltAt: string | undefined;
    for (let i = 0; i < fileRefs.length; i += CHUNK) {
      const chunk = fileRefs.slice(i, i + CHUNK);
      const res = await this._post<IBacklinksBatchResponse>('/api/link-inventory/backlinks/batch', { fileRefs: chunk });
      mergedAdmin = mergedAdmin || res.isAdmin;
      if (res.indexBuiltAt) mergedBuiltAt = res.indexBuiltAt;
      Object.assign(merged.results, res.results);
    }
    merged.isAdmin = mergedAdmin;
    merged.indexBuiltAt = mergedBuiltAt;
    return merged;
  }

  public async getDuplicatesBootstrapStatus(jobId: string): Promise<IDuplicatesBootstrapJob> {
    const res = await this._get<{ ok: boolean; job: IDuplicatesBootstrapJob }>(`/api/duplicates/bootstrap/${jobId}`);
    return res.job;
  }
}

export interface IBacklinkSource {
  sourceKind: 'page' | 'document';
  site: string;
  siteUrl: string;
  title: string;
  url: string;
  scannedAt: string;
}

export interface IBacklinksResultEntry {
  visible: IBacklinkSource[];
  hiddenCount: number;
}

export interface IBacklinksBatchResponse {
  indexBuiltAt?: string;
  isAdmin: boolean;
  results: Record<string, IBacklinksResultEntry>;
}

export interface IDuplicatesBootstrapStartResponse {
  ok: boolean;
  jobId?: string;
  status?: 'queued';
  sitePath?: string;
  libraryTitle?: string;
  filesTotal?: number;
  maxVersionsPerFile?: number;
  info?: string;
}

export interface IDuplicatesBootstrapJob {
  jobId: string;
  status: 'queued' | 'running' | 'completed' | 'failed';
  startedAt: string;
  finishedAt?: string;
  sitePath: string;
  libraryTitle: string;
  maxVersionsPerFile: number;
  includeOther: boolean;
  filesTotal: number;
  filesCompleted: number;
  versionsProcessed: number;
  errorCount: number;
  recentErrors: string[];
  currentFile?: string;
  caller?: string;
}

export interface IDuplicatesAllowlistEntry {
  note: string;
  addedBy: string;
  addedAt: string;
}

export interface IDuplicatesAllowlist {
  hashAllowlist: Array<IDuplicatesAllowlistEntry & { sha256: string }>;
  pathAllowlist: Array<IDuplicatesAllowlistEntry & { pattern: string }>;
  nameAllowlist: Array<IDuplicatesAllowlistEntry & { pattern: string }>;
}

export interface IDuplicatesExactGroup {
  sha256: string;
  files: Array<{
    fileRef: string;
    fileName: string;
    sitePath: string;
    library: string;
    size: number;
    currentHashObservedAt: string;
  }>;
}

export interface IDuplicatesStalePair {
  staleFileRef: string;
  staleFileName: string;
  staleSitePath: string;
  staleLibrary: string;
  staleCurrentHash: string;
  authoritativeFileRef: string;
  authoritativeFileName: string;
  authoritativeSitePath: string;
  authoritativeLibrary: string;
  authoritativeCurrentHash: string;
  divergedAt: string;
}

export interface IDuplicatesSameNamePair {
  aFileRef: string;
  aFileName: string;
  aSitePath: string;
  aSize: number;
  aSha256: string;
  bFileRef: string;
  bFileName: string;
  bSitePath: string;
  bSize: number;
  bSha256: string;
  sameSize: boolean;
}

/** Two files whose normalized text is similar (SimHash Hamming distance ≤ 3 by default). */
export interface IDuplicatesNearDuplicatePair {
  aFileRef: string;
  aFileName: string;
  aSitePath: string;
  aLibrary: string;
  aSize: number;
  aSimhash: string;
  bFileRef: string;
  bFileName: string;
  bSitePath: string;
  bLibrary: string;
  bSize: number;
  bSimhash: string;
  hammingDistance: number;
}

/**
 * Two files whose currents differ but whose version histories share an
 * ancestor — both copies forked from a common origin.
 */
export interface IDuplicatesDivergedPair {
  aFileRef: string;
  aFileName: string;
  aSitePath: string;
  aLibrary: string;
  aCurrentHash: string;
  bFileRef: string;
  bFileName: string;
  bSitePath: string;
  bLibrary: string;
  bCurrentHash: string;
  sharedAncestorTextHash: string;
  ancestorObservedAt: string;
}

export interface IDuplicatesReport {
  indexBuiltAt: string | null;
  isAdmin: boolean;
  totals: {
    exactGroups: number;
    staleFiles: number;
    sameNamePairs: number;
    nearDuplicatePairs: number;
    divergedPairs: number;
  };
  exactGroups: IDuplicatesExactGroup[];
  stalePairs: IDuplicatesStalePair[];
  sameNamePairs: IDuplicatesSameNamePair[];
  nearDuplicatePairs: IDuplicatesNearDuplicatePair[];
  divergedPairs: IDuplicatesDivergedPair[];
  truncated?: {
    exactGroups: boolean;
    stalePairs: boolean;
    sameNamePairs: boolean;
    nearDuplicatePairs: boolean;
    divergedPairs: boolean;
  };
}

export interface IDuplicatesLookupResponse {
  fileRef: string;
  sha256?: string;
  known: boolean;
  isAdmin: boolean;
  indexBuiltAt?: string;
  exact: Array<{
    fileRef: string;
    fileName: string;
    sitePath: string;
    library: string;
    size: number;
    sha256: string;
    currentHashObservedAt: string;
  }>;
  stale: Array<{
    authoritativeFileRef: string;
    authoritativeFileName: string;
    authoritativeSitePath: string;
    authoritativeLibrary: string;
    authoritativeCurrentHash: string;
    divergedAt: string;
    matchingPreviousHash: string;
  }>;
  sameName: Array<{
    fileRef: string;
    fileName: string;
    sitePath: string;
    library: string;
    size: number;
    sha256: string;
    sameSize: boolean;
  }>;
  hiddenCount: { exact: number; stale: number; sameName: number };
}

export interface IEnableBacklinksColumnResponse {
  ok: boolean;
  siteUrl: string;
  libraryTitle: string;
  columnInternalName: string;
  columnDisplayName: string;
  customizerRegistered: boolean;
  addedToDefaultView: boolean;
}

export interface IRebuildBacklinksResponse {
  ok: boolean;
  builtAt: string;
  sitesTotal: number;
  sitesWithPageData: number;
  sitesWithDocData: number;
  sourcePageJobId?: string;
  sourceDocJobId?: string;
}

export interface IReplaceDetail {
  oldUrl: string;
  newUrl: string;
  matched: boolean;
}

export interface IReplacePageResult {
  site: string;
  pageId: number;
  pageTitle?: string;
  /** Server-relative URL of the patched page, when known. Used by the
   *  result row to deep-link directly to the page that was modified. */
  pageUrl?: string;
  status: 'applied' | 'preview' | 'stale' | 'conflict' | 'no-match' | 'not-found' | 'error';
  replacementCount: number;
  unmatchedCount: number;
  details: IReplaceDetail[];
  error?: string;
  newEtag?: string;
}

export interface ILinkInventoryReplaceResponse {
  ok: boolean;
  dryRun: boolean;
  summary: {
    totalPages: number;
    applied: number;
    preview: number;
    stale: number;
    conflict: number;
    noMatch: number;
    error: number;
  };
  droppedSites?: string[];
  results: IReplacePageResult[];
}
