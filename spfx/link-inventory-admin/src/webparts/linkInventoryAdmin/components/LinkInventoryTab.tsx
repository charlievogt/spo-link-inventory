import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, SelectionMode } from '@fluentui/react/lib/DetailsList';
import type { IColumn } from '@fluentui/react/lib/DetailsList';
import { CommandBar, type ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { Dropdown, type IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Stack } from '@fluentui/react/lib/Stack';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Text } from '@fluentui/react/lib/Text';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import {
  type LinkInventoryService,
  type ILinkInventoryJob,
  type ILinkInventoryResults,
  type IClassifiedLink,
  type IPageInventory,
  type ISiteInventory,
  type ILinkInventoryReplaceResponse,
  type IReplacePageResult,
  type ISiteSummary,
  type IHubSummary,
  type IPreviewResults,
  type IPreviewBucketStats,
  type IEnableBacklinksColumnResponse,
  type IRebuildBacklinksResponse,
  type IScheduleConfig,
} from '../services/LinkInventoryService';
import { TextField } from '@fluentui/react/lib/TextField';
import { clientSuggestReplacement, clientCanonicalKey, clientMalformedCheck, TENANT_ORIGIN } from './linkSuggestionClient';

/**
 * Link Inventory tab.
 *
 * Lets admins trigger tenant-wide page scans and lets anyone with read
 * access to a scanned site browse the resulting link inventory: every
 * URL on every modern page across the in-scope sites, with filters by
 * link class (onprem / sharing-link / external / etc.), site, web part
 * source, and free-text search.
 *
 * Admin gating:
 *   - Triggering a scan requires membership in the SP Redirect Manager
 *     Admins Entra group. Backend enforces (POST returns 403); the UI
 *     hides the button via the whoami probe so non-admins don't see a
 *     button that won't work.
 *
 * Per-site read filtering:
 *   - The results endpoint returns only sites the calling user can read.
 *     Sites the user lacks access to are surfaced as `droppedSites` so
 *     the UI can tell them "8 of 12 sites visible to you".
 */

interface ILinkInventoryTabProps {
  service: LinkInventoryService;
  onError: (msg: string) => void;
}

interface ILinkRow {
  /** Composite key — jobId + page url + link index — used by DetailsList. */
  key: string;
  site: string;
  page: IPageInventory;
  link: IClassifiedLink;
  /** Source job id — present so the detail panel can render which scan
   *  a sibling occurrence came from. */
  jobId: string;
  /**
   * Source job kind — drives whether write-back actions are available
   * and how the detail panel partitions sibling rows. With unified
   * scans this is set per-row from `page.sourceKind`, so a single
   * inventory can carry rows from both kinds.
   */
  jobKind: 'pages' | 'documents';
}

/**
 * Map from canonical key to the list of rows that share it. Built at
 * results-load time over BOTH the selected job and its companion job
 * (the most recent of the opposite kind), so the detail panel can show
 * "also linked from these document files" while viewing a page scan
 * (and vice versa).
 */
type CanonicalIndex = Map<string, ILinkRow[]>;

/**
 * One entry in the "Errors" panel — either a site that failed to scan
 * entirely (`kind: 'site'`) or a single page/file that couldn't be
 * parsed (`kind: 'page'`). Errored items never appear in the link
 * table because they have 0 links, so this is the only place a user
 * can see what went wrong.
 */
interface IScanErrorRow {
  site: string;
  kind: 'site' | 'page';
  /** Display label — page title for pages, site path for site errors. */
  target: string;
  /** Server-relative path of the file/page (page errors only). */
  url?: string;
  error: string;
}

/**
 * One row in the Canonicalization Report — a non-canonical link from
 * a page or document file pointing at a SPO document, with the
 * suggested canonical AllItems URL alongside it.
 */
interface ICanonicalReportRow {
  /** Source page or doc file containing the link. */
  source: string;
  /** Server-relative URL of the source. */
  sourceUrl: string;
  /** Site path of the source. */
  site: string;
  /** Whether the source is a page or a document file. */
  sourceKind: 'page' | 'document';
  /** Visible link text (or the target file name if no text). */
  text: string;
  /** Current URL as it appears in the source. */
  current: string;
  /** Suggested canonical AllItems URL. */
  suggested: string;
  /** Link class — informational. */
  linkClass: string;
}

interface ILinkInventoryTabState {
  loading: boolean;
  isAdmin: boolean;
  upn: string;
  jobs: ILinkInventoryJob[];
  selectedJobId: string | undefined;
  selectedJob: ILinkInventoryJob | undefined;
  results: ILinkInventoryResults | undefined;
  /** Populated instead of `results` when the selected job is a preview-only
   *  doc scan — shape is per-site/per-library counts, not per-page links. */
  previewResults: IPreviewResults | undefined;
  /** Sites the user has ticked for promotion in the preview view.
   *  Empty = all sites in the preview will be promoted. */
  previewSelectedSites: Set<string>;
  /** True while a promote-to-scan call is in flight. */
  promoting: boolean;
  /**
   * Index over results — canonical key → all rows that share it. For
   * unified scans, the merged inventory in `results` already carries
   * both kinds (each page tagged with `sourceKind`), so this index
   * spans pages and docs in a single map.
   */
  canonicalIndex: CanonicalIndex;
  resultsLoading: boolean;
  triggering: boolean;
  polling: boolean;
  /** Page-scan options dialog open */
  pageScanOpen: boolean;
  /** Set of selected server-relative site paths for the next page scan. Empty = full tenant. */
  pageScanSelectedSites: Set<string>;
  /** Filter text for the page-scan site picker. */
  pageScanSitePickerFilter: string;
  /** Opt-in SPO file verification for the next page scan. */
  pageScanVerifyFiles: boolean;
  /** Doc-scan options dialog open */
  docScanOpen: boolean;
  /** Max file size in MB for the next document scan (default 100). */
  docScanMaxMB: number;
  /** Set of selected server-relative site paths for the next doc scan. Empty = full tenant. */
  docScanSelectedSites: Set<string>;
  /** Newline-separated list of file paths for a targeted doc scan (overrides sites). */
  docScanFileRefs: string;
  /** Modified-after date filter (YYYY-MM-DD) for incremental rescans. Empty = no filter. */
  docScanModifiedAfter: string;
  /** Preview-only mode — enumerate files and report counts without scanning. */
  docScanPreviewOnly: boolean;
  /** Opt-in SPO file verification for the next doc scan. */
  docScanVerifyFiles: boolean;
  /** Filter text for the site picker. */
  docScanSitePickerFilter: string;
  /** Loaded tenant site list for the picker. Lazily fetched when the panel opens. */
  sitePickerSites: ISiteSummary[];
  sitePickerHubs: IHubSummary[];
  sitePickerLoading: boolean;
  sitePickerError: string;
  // filters
  filterClass: string;
  filterSite: string;
  filterSource: string;
  searchText: string;
  // detail panel
  detailRow: ILinkRow | undefined;
  /** Bumped on panel dismiss to force the DetailsList to remount and
   *  clear its internal "active item" state. Without this, Fluent UI
   *  re-fires onActiveItemChanged when focus returns to the list and
   *  the panel pops back open immediately. */
  detailsListKey: number;
  /** Whether the "hidden sites" list is expanded (default collapsed). */
  droppedSitesExpanded: boolean;
  /** Whether the "errored files/pages" list is expanded (default collapsed). */
  errorsExpanded: boolean;
  /** Confirmation dialog for the Delete this scan action. */
  deleteConfirmOpen: boolean;
  /** True while a delete API call is in flight. */
  deleting: boolean;
  /** Canonicalization report panel open. */
  canonicalReportOpen: boolean;
  // find/replace
  findReplaceOpen: boolean;
  frFind: string;
  frReplace: string;
  /** Match mode for find/replace. 'canonical' is set by Align All. */
  frMode: 'substring' | 'canonical';
  /** Optional human label for the mode (shown in the panel header). */
  frModeLabel: string;
  /** In-flight call to /replace */
  frBusy: boolean;
  /** Server response from the most recent dry-run / apply */
  frResponse: ILinkInventoryReplaceResponse | undefined;
  /** True after a successful real-run, so we show the success state */
  frApplied: boolean;
  /** Error from the most recent call */
  frError: string;

  // Backlinks column management
  backlinksDialogOpen: boolean;
  backlinksSiteUrl: string;
  backlinksLibraryTitle: string;
  backlinksBusy: boolean;
  backlinksError: string;
  backlinksEnableResult: IEnableBacklinksColumnResponse | undefined;
  rebuildingBacklinks: boolean;
  rebuildResult: IRebuildBacklinksResponse | undefined;
  rebuildError: string;
  exportingBacklinks: boolean;

  // Daily-scan schedule (Schedule panel)
  scheduleOpen: boolean;
  scheduleConfig: IScheduleConfig | undefined;
  scheduleLoading: boolean;
  scheduleSaving: boolean;
  scheduleError: string;
  /** Draft values mirrored from scheduleConfig while the panel is open. */
  scheduleEnabledDraft: boolean;
  scheduleTimeOfDayDraft: string;
  scheduleTimeZoneDraft: string;
}

const LINK_CLASS_OPTIONS: IDropdownOption[] = [
  { key: '', text: 'All classes' },
  { key: 'malformed-spo-link', text: 'Malformed SPO link' },
  { key: 'onprem', text: 'On-prem' },
  { key: 'sharing-link', text: 'Sharing link' },
  { key: 'office-online', text: 'Office Online wrapper' },
  { key: 'doc-aspx', text: 'Doc.aspx (sourcedoc)' },
  { key: 'spo-internal', text: 'SPO internal (absolute)' },
  { key: 'relative', text: 'Relative path' },
  { key: 'external', text: 'External' },
  { key: 'mailto', text: 'mailto:' },
  { key: 'tel', text: 'tel:' },
  { key: 'anchor-only', text: 'Anchor (#fragment)' },
  { key: 'javascript', text: 'javascript:' },
  { key: 'unknown', text: 'Unknown' },
];

const MALFORMED_REASON_LABEL: Record<string, string> = {
  'search-fragment': 'Leftover #search= fragment — URL was copied from a library-search session. Often a sign of a bad find/replace.',
  'file-not-found': 'File at ?id= returned 404 — the referenced file no longer exists at that path.',
  'both': 'Leftover #search= fragment AND the referenced file returned 404.',
};

export class LinkInventoryTab extends React.Component<ILinkInventoryTabProps, ILinkInventoryTabState> {
  private pollTimer: number | undefined;

  public constructor(props: ILinkInventoryTabProps) {
    super(props);
    this.state = {
      loading: true,
      isAdmin: false,
      upn: '',
      jobs: [],
      selectedJobId: undefined,
      selectedJob: undefined,
      results: undefined,
      previewResults: undefined,
      previewSelectedSites: new Set(),
      promoting: false,
      canonicalIndex: new Map(),
      resultsLoading: false,
      triggering: false,
      polling: false,
      pageScanOpen: false,
      pageScanSelectedSites: new Set(),
      pageScanSitePickerFilter: '',
      pageScanVerifyFiles: false,
      docScanOpen: false,
      docScanMaxMB: 100,
      docScanSelectedSites: new Set(),
      docScanFileRefs: '',
      docScanModifiedAfter: '',
      docScanPreviewOnly: true,
      docScanVerifyFiles: false,
      docScanSitePickerFilter: '',
      sitePickerSites: [],
      sitePickerHubs: [],
      sitePickerLoading: false,
      sitePickerError: '',
      filterClass: '',
      filterSite: '',
      filterSource: '',
      searchText: '',
      detailRow: undefined,
      detailsListKey: 0,
      droppedSitesExpanded: false,
      errorsExpanded: false,
      deleteConfirmOpen: false,
      deleting: false,
      canonicalReportOpen: false,
      findReplaceOpen: false,
      frFind: '',
      frReplace: '',
      frMode: 'substring',
      frModeLabel: '',
      frBusy: false,
      frResponse: undefined,
      frApplied: false,
      frError: '',
      backlinksDialogOpen: false,
      backlinksSiteUrl: '',
      backlinksLibraryTitle: 'Shared Documents',
      backlinksBusy: false,
      backlinksError: '',
      backlinksEnableResult: undefined,
      rebuildingBacklinks: false,
      rebuildResult: undefined,
      rebuildError: '',
      exportingBacklinks: false,
      scheduleOpen: false,
      scheduleConfig: undefined,
      scheduleLoading: false,
      scheduleSaving: false,
      scheduleError: '',
      scheduleEnabledDraft: false,
      scheduleTimeOfDayDraft: '04:00',
      scheduleTimeZoneDraft: 'America/Chicago',
    };
  }

  public componentDidMount(): void {
    this._init().catch((err: Error) => this.props.onError(err.message));
  }

  public componentWillUnmount(): void {
    this._stopPolling();
  }

  private async _init(): Promise<void> {
    try {
      const [whoami, jobs] = await Promise.all([
        this.props.service.whoami(),
        this.props.service.listJobs(),
      ]);
      // Default selection: most recent completed job with results, otherwise the most recent job at all.
      const completed = jobs.find((j) => j.resultsAvailable);
      const selectedJobId = completed?.jobId ?? jobs[0]?.jobId;
      this.setState({
        loading: false,
        isAdmin: whoami.isAdmin,
        upn: whoami.upn ?? whoami.userId,
        jobs,
        selectedJobId,
        selectedJob: jobs.find((j) => j.jobId === selectedJobId),
      });
      if (selectedJobId) {
        const job = jobs.find((j) => j.jobId === selectedJobId);
        if (job?.resultsAvailable) await this._loadResults(selectedJobId);
        else if (job && (job.status === 'queued' || job.status === 'running')) this._startPolling(selectedJobId);
      }
    } catch (err) {
      this.setState({ loading: false });
      throw err;
    }
  }

  private async _loadResults(jobId: string): Promise<void> {
    this.setState({
      resultsLoading: true,
      results: undefined,
      previewResults: undefined,
      previewSelectedSites: new Set(),
      canonicalIndex: new Map(),
    });

    // If the selected job is a preview-only doc scan, load the preview
    // results blob (per-site counts) instead of the regular link inventory.
    const job = this.state.jobs.find((j) => j.jobId === jobId);
    if (job?.kind === 'documents' && job.previewOnly) {
      try {
        const previewResults = await this.props.service.getPreviewResults(jobId);
        this.setState({
          previewResults,
          previewSelectedSites: new Set(previewResults.siteSummaries.map((s) => s.site)),
          resultsLoading: false,
        });
      } catch (err) {
        this.setState({ resultsLoading: false });
        this.props.onError(`Loading preview results: ${(err as Error).message}`);
      }
      return;
    }

    try {
      const primaryResults = await this.props.service.getResults(jobId);
      this._enrichLinks(primaryResults);

      // For unified scans, fetch the sibling job's results and MERGE
      // them into a single inventory so the table, totals, and canonical
      // index all see the union. Each page is tagged with its
      // `sourceKind` during the merge so per-row rendering can show
      // which scan kind it came from.
      const primaryJob = this.state.jobs.find((j) => j.jobId === jobId);
      const primaryKind: 'pages' | 'documents' = primaryJob?.kind ?? 'pages';
      const oppositeKind: 'pages' | 'documents' = primaryKind === 'documents' ? 'pages' : 'documents';
      const companionJob = this.state.jobs.find(
        (j) => (j.kind ?? 'pages') === oppositeKind && j.resultsAvailable && j.jobId !== jobId,
      );

      let merged: ILinkInventoryResults = this._tagSourceKind(primaryResults, primaryKind);

      if (companionJob) {
        try {
          const companionResults = await this.props.service.getResults(companionJob.jobId);
          this._enrichLinks(companionResults);
          merged = this._mergeInventories(merged, this._tagSourceKind(companionResults, oppositeKind));
        } catch (err) {
          // Non-fatal — main view still renders without the companion.
          // eslint-disable-next-line no-console
          console.warn(`Companion job load failed (${companionJob.jobId}): ${(err as Error).message}`);
        }
      }

      const canonicalIndex = this._buildCanonicalIndex(merged, jobId);
      this.setState({
        results: merged,
        canonicalIndex,
        resultsLoading: false,
      });
    } catch (err) {
      this.setState({ resultsLoading: false });
      this.props.onError(`Loading results: ${(err as Error).message}`);
    }
  }

  /**
   * Backfill `suggestion` and `canonicalKey` on links from older scan
   * blobs that pre-date the server-side suggestion feature. Mutates
   * each link object in place. Newer scans already have these fields
   * populated server-side; we never overwrite an existing value.
   *
   * Without this, the "Fix this link", "Align all to canonical", and
   * sibling-grouping features only work for fresh scans.
   */
  private _enrichLinks(results: ILinkInventoryResults): void {
    for (const site of results.sites) {
      // Track whether any link's class was upgraded by the client
      // heuristic so we can re-tally byClass at the end.
      let mutated = false;
      for (const page of site.pages) {
        for (const link of page.links) {
          // Upgrade class FIRST so the suggestion computation below
          // sees the post-upgrade class. Without this, old blobs that
          // have `#search=` AllItems URLs would stay `spo-internal`
          // through the suggestion step and never get the fragment-
          // strip suggestion assigned.
          if (link.linkClass !== 'malformed-spo-link') {
            const mf = clientMalformedCheck(link.url);
            if (mf) {
              link.linkClass = mf.linkClass;
              link.malformedReason = mf.malformedReason;
              mutated = true;
            }
          }
          if (link.suggestion === undefined) {
            const s = clientSuggestReplacement(link.url, link.linkClass);
            if (s) link.suggestion = s;
          }
          if (!link.canonicalKey) {
            const k = clientCanonicalKey(link.url, link.linkClass, link.suggestion ?? '');
            if (k) link.canonicalKey = k;
          }
        }
      }
      if (mutated) {
        const byClass: Record<string, number> = {};
        for (const page of site.pages) {
          for (const link of page.links) {
            byClass[link.linkClass] = (byClass[link.linkClass] ?? 0) + 1;
          }
        }
        site.byClass = byClass;
      }
    }
  }

  /**
   * Build a row from one link, with stable composite key.
   *
   * Composite key includes jobId so cross-job siblings deduplicate
   * correctly, and pageUrl so multiple files in a doc scan (which all
   * share pageId=0) don't collide.
   *
   * `jobKind` is taken from the page's `sourceKind` tag (set during the
   * unified-scan merge) when present, falling back to the selected
   * job's kind otherwise.
   */
  private _makeRow(
    site: string,
    page: IPageInventory,
    link: IClassifiedLink,
    linkIdx: number,
    jobId: string,
    jobKind: 'pages' | 'documents',
  ): ILinkRow {
    return {
      key: `${jobId}:${page.pageUrl}:${linkIdx}`,
      site,
      page,
      link,
      jobId,
      jobKind,
    };
  }

  /**
   * Tag every page in a results blob with its `sourceKind`. Used at
   * merge time to remember which scan kind each page/file came from.
   * Mutates in place — the returned reference is the same object.
   */
  private _tagSourceKind(
    results: ILinkInventoryResults,
    kind: 'pages' | 'documents',
  ): ILinkInventoryResults {
    for (const site of results.sites) {
      for (const page of site.pages) {
        page.sourceKind = kind;
      }
    }
    return results;
  }

  /**
   * Merge a companion inventory into a primary inventory. Sites that
   * appear in both have their `pages[]` concatenated and counters
   * summed. Per-page `sourceKind` is preserved from `_tagSourceKind`,
   * so the merged inventory knows which rows came from which scan.
   *
   * Totals: `pages` and `links` are summed (a page-found page and a
   * doc-found file are both "things scanned"); `sites` uses the union
   * count rather than naive sum because the same site path can appear
   * in both inventories.
   */
  private _mergeInventories(
    primary: ILinkInventoryResults,
    companion: ILinkInventoryResults,
  ): ILinkInventoryResults {
    const sitesByPath = new Map<string, ISiteInventory>();
    for (const site of primary.sites) sitesByPath.set(site.site, site);
    for (const cSite of companion.sites) {
      const existing = sitesByPath.get(cSite.site);
      if (existing) {
        existing.pages = [...existing.pages, ...cSite.pages];
        existing.pageCount += cSite.pageCount;
        existing.linkCount += cSite.linkCount;
        existing.scanMs += cSite.scanMs;
        for (const k of Object.keys(cSite.byClass)) {
          existing.byClass[k] = (existing.byClass[k] ?? 0) + cSite.byClass[k];
        }
        for (const k of Object.keys(cSite.bySource)) {
          existing.bySource[k] = (existing.bySource[k] ?? 0) + cSite.bySource[k];
        }
      } else {
        sitesByPath.set(cSite.site, cSite);
      }
    }

    const earlier = (a: string, b: string): string => (a < b ? a : b);
    const later = (a: string, b: string): string => (a > b ? a : b);
    const droppedSet = new Set<string>([...primary.droppedSites, ...companion.droppedSites]);

    return {
      jobId: primary.jobId,
      startedAt: earlier(primary.startedAt, companion.startedAt),
      finishedAt: later(primary.finishedAt, companion.finishedAt),
      sites: Array.from(sitesByPath.values()),
      totals: {
        sites: sitesByPath.size,
        pages: primary.totals.pages + companion.totals.pages,
        links: primary.totals.links + companion.totals.links,
        errors: primary.totals.errors + companion.totals.errors,
      },
      droppedSites: Array.from(droppedSet),
      userScopedTotals: {
        sites: Math.max(primary.userScopedTotals.sites, companion.userScopedTotals.sites),
        pages: primary.userScopedTotals.pages + companion.userScopedTotals.pages,
        links: primary.userScopedTotals.links + companion.userScopedTotals.links,
        errors: primary.userScopedTotals.errors + companion.userScopedTotals.errors,
      },
    };
  }

  /**
   * Walk every link in the merged inventory and build the canonical-key
   * → rows index. Empty canonical keys are skipped (links that can't be
   * canonicalized — opaque sharing tokens, mailto, etc. — never group
   * with anything).
   *
   * Each row's `jobKind` comes from its page's `sourceKind` so both
   * page-found and doc-found rows are present in the index, and the
   * detail panel can show "also linked from N pages, M files" using the
   * jobKind comparison.
   */
  private _buildCanonicalIndex(
    results: ILinkInventoryResults,
    jobId: string,
  ): CanonicalIndex {
    const idx: CanonicalIndex = new Map();
    this._addToIndex(idx, results, jobId);
    return idx;
  }

  private _addToIndex(
    idx: CanonicalIndex,
    results: ILinkInventoryResults,
    jobId: string,
  ): void {
    for (const site of results.sites) {
      for (const page of site.pages) {
        const pageKind: 'pages' | 'documents' = page.sourceKind ?? 'pages';
        for (let i = 0; i < page.links.length; i++) {
          const link = page.links[i];
          const key = link.canonicalKey;
          if (!key) continue;
          const row = this._makeRow(site.site, page, link, i, jobId, pageKind);
          if (!idx.has(key)) idx.set(key, []);
          idx.get(key)!.push(row);
        }
      }
    }
  }

  private _startPolling(jobId: string): void {
    this._stopPolling();
    this.setState({ polling: true });
    const tick = async (): Promise<void> => {
      try {
        const job = await this.props.service.getJob(jobId);
        let updatedJobs = this.state.jobs.map((j) => (j.jobId === jobId ? job : j));
        // If this job is part of a unified pair, also refresh the sibling
        // so the combined dropdown label and any combined progress UI
        // stays accurate.
        let sibling: ILinkInventoryJob | undefined;
        if (job.siblingJobId) {
          try {
            sibling = await this.props.service.getJob(job.siblingJobId);
            updatedJobs = updatedJobs.map((j) => (j.jobId === sibling!.jobId ? sibling! : j));
          } catch {
            // Sibling fetch is best-effort; main job state still updates.
          }
        }
        this.setState({ jobs: updatedJobs, selectedJob: job });
        const primaryDone = job.status === 'completed' || job.status === 'failed';
        const siblingDone = !sibling || sibling.status === 'completed' || sibling.status === 'failed';
        if (primaryDone && siblingDone) {
          this._stopPolling();
          if (job.resultsAvailable) await this._loadResults(jobId);
        }
      } catch (err) {
        this._stopPolling();
        this.props.onError(`Polling job: ${(err as Error).message}`);
      }
    };
    this.pollTimer = window.setInterval(() => { void tick(); }, 2000);
    void tick();
  }

  private _stopPolling(): void {
    if (this.pollTimer) {
      window.clearInterval(this.pollTimer);
      this.pollTimer = undefined;
    }
    this.setState({ polling: false });
  }

  /**
   * Trigger a unified scan — pages + docs in one go. Defaults: tenant-wide
   * (no site picker), default file size cap, no preview, no per-link
   * verifier. The two halves are sibling-linked; the polling loop
   * follows both. We select the page job for the dropdown since pages
   * support find-and-replace.
   */
  private _triggerUnifiedScan = async (): Promise<void> => {
    this.setState({ triggering: true });
    try {
      const res = await this.props.service.triggerUnifiedScan({});
      const jobs = await this.props.service.listJobs();
      this.setState({
        jobs,
        selectedJobId: res.pageJobId,
        selectedJob: jobs.find((j) => j.jobId === res.pageJobId),
        results: undefined,
        triggering: false,
      });
      this._startPolling(res.pageJobId);
    } catch (err) {
      this.setState({ triggering: false });
      this.props.onError(`Trigger unified scan: ${(err as Error).message}`);
    }
  };

  private _triggerScan = async (): Promise<void> => {
    this.setState({ triggering: true, pageScanOpen: false });
    try {
      // Build the optional targeted sites list from the picker. Empty
      // selection = full tenant scan.
      const sites = Array.from(this.state.pageScanSelectedSites);
      const opts: { sites?: string[]; verifyFiles?: boolean } = {};
      if (sites.length > 0) opts.sites = sites;
      if (this.state.pageScanVerifyFiles) opts.verifyFiles = true;
      const res = await this.props.service.triggerScan(opts);
      // Refresh job list and select the new one
      const jobs = await this.props.service.listJobs();
      this.setState({
        jobs,
        selectedJobId: res.jobId,
        selectedJob: jobs.find((j) => j.jobId === res.jobId),
        results: undefined,
        triggering: false,
      });
      this._startPolling(res.jobId);
    } catch (err) {
      this.setState({ triggering: false });
      this.props.onError(`Trigger scan: ${(err as Error).message}`);
    }
  };

  private _triggerDocScan = async (): Promise<void> => {
    this.setState({ triggering: true, docScanOpen: false });
    try {
      const maxFileBytes = Math.max(1, Math.floor(this.state.docScanMaxMB)) * 1024 * 1024;
      const sites = Array.from(this.state.docScanSelectedSites);
      const fileRefs = this.state.docScanFileRefs
        .split(/\r?\n/)
        .map((s) => s.trim())
        .filter((s) => s.length > 0);
      const opts: {
        maxFileBytes: number;
        sites?: string[];
        fileRefs?: string[];
        modifiedAfter?: string;
        previewOnly?: boolean;
        verifyFiles?: boolean;
      } = { maxFileBytes };
      if (this.state.docScanVerifyFiles) opts.verifyFiles = true;
      if (fileRefs.length > 0) {
        opts.fileRefs = fileRefs;
      } else if (sites.length > 0) {
        opts.sites = sites;
      }
      const modifiedAfter = this.state.docScanModifiedAfter.trim();
      if (modifiedAfter) {
        // <input type="date"> gives YYYY-MM-DD; convert to ISO with explicit UTC.
        opts.modifiedAfter = new Date(`${modifiedAfter}T00:00:00Z`).toISOString();
      }
      if (this.state.docScanPreviewOnly && fileRefs.length === 0) {
        opts.previewOnly = true;
      }
      const res = await this.props.service.triggerDocScan(opts);
      if (!res.jobId) {
        // No files were enumerated — surface as a friendly message, not an error
        this.props.onError(res.message ?? 'Document scan returned no files.');
        this.setState({ triggering: false });
        return;
      }
      const jobs = await this.props.service.listJobs();
      this.setState({
        jobs,
        selectedJobId: res.jobId,
        selectedJob: jobs.find((j) => j.jobId === res.jobId),
        results: undefined,
        triggering: false,
      });
      this._startPolling(res.jobId);
    } catch (err) {
      this.setState({ triggering: false });
      this.props.onError(`Trigger document scan: ${(err as Error).message}`);
    }
  };

  private _refresh = async (): Promise<void> => {
    try {
      const jobs = await this.props.service.listJobs();
      this.setState({ jobs, selectedJob: jobs.find((j) => j.jobId === this.state.selectedJobId) });
      if (this.state.selectedJobId && this.state.selectedJob?.resultsAvailable) {
        await this._loadResults(this.state.selectedJobId);
      }
    } catch (err) {
      this.props.onError(`Refresh: ${(err as Error).message}`);
    }
  };

  private _onJobChange = (_e: unknown, opt?: IDropdownOption): void => {
    if (!opt) return;
    const jobId = String(opt.key);
    const job = this.state.jobs.find((j) => j.jobId === jobId);
    this.setState({ selectedJobId: jobId, selectedJob: job, results: undefined });
    this._stopPolling();
    if (job?.resultsAvailable) void this._loadResults(jobId);
    else if (job && (job.status === 'queued' || job.status === 'running')) this._startPolling(jobId);
  };

  /**
   * Walk the currently-selected scan's results and collect every
   * non-canonical link to a SPO document. A link is considered
   * non-canonical when it's a sharing-link wrapper, an Office Online
   * wrapper, or a direct SPO file URL that isn't already in
   * `Forms/AllItems.aspx?id=…` form. Each row carries the source
   * (page or document file), the link text, the current URL, and
   * the suggested canonical AllItems URL.
   *
   * Scoped to the **primary job only** — the cross-job companion
   * scan is intentionally NOT walked because the user expects the
   * report to reflect what was scanned, not the union of two scans.
   * If a future requirement needs cross-scan rollup, add a toggle
   * to the panel rather than making it implicit.
   *
   * Used by the Canonicalization Report panel.
   */
  private _collectCanonicalReport(): ICanonicalReportRow[] {
    const results = this.state.results;
    if (!results) return [];
    const sourceKind: 'page' | 'document' =
      this.state.selectedJob?.kind === 'documents' ? 'document' : 'page';

    const out: ICanonicalReportRow[] = [];
    const seen = new Set<string>(); // dedupe key: source + current URL

    for (const site of results.sites) {
      for (const page of site.pages) {
        for (const link of page.links) {
          // Qualifying classes: sharing-link / office-online (always
          // rewritten to AllItems), non-canonical spo-internal (direct
          // file URL rewritten to AllItems), and malformed-spo-link
          // with a suggestion (search-fragment case — fragment stripped).
          // The suggestion-presence check below catches the cases where
          // a class would otherwise qualify but no safe rewrite exists
          // (already canonical spo-internal; malformed-spo-link with
          // file-not-found reason, since the verifier cleared it).
          const cls = link.linkClass;
          if (
            cls !== 'sharing-link' &&
            cls !== 'office-online' &&
            cls !== 'spo-internal' &&
            cls !== 'malformed-spo-link'
          ) {
            continue;
          }
          if (!link.suggestion || link.suggestion.length === 0) continue;

          // Dedupe — the same (source page/file, target URL) pair
          // can show up multiple times when a single page contains
          // duplicate links. The user only needs to fix it once.
          const dedupeKey = `${page.pageUrl}||${link.url}`;
          if (seen.has(dedupeKey)) continue;
          seen.add(dedupeKey);

          out.push({
            source: page.pageTitle || page.pageUrl,
            sourceUrl: page.pageUrl,
            site: site.site,
            sourceKind,
            text: link.text ?? '',
            current: link.url,
            suggested: link.suggestion,
            linkClass: cls,
          });
        }
      }
    }

    // Stable sort by source then by current URL so re-runs give the
    // same row order.
    out.sort((a, b) => {
      const s = a.source.localeCompare(b.source);
      if (s !== 0) return s;
      return a.current.localeCompare(b.current);
    });

    return out;
  }

  /**
   * CSV export of the canonicalization report. The columns are
   * exactly what the user asked for: file/page, link text, current
   * link, suggested canonical (viewer) link. Site and source kind
   * are appended as bonus columns for downstream filtering.
   */
  private _exportCanonicalReportCsv = (): void => {
    const rows = this._collectCanonicalReport();
    const header = ['File/Page', 'Site', 'Source kind', 'Link text', 'Current link', 'Suggested canonical link', 'Link class'];
    const escape = (s: string): string => {
      if (s == null) return '';
      const needsQuotes = /[",\r\n]/.test(s);
      const escaped = s.replace(/"/g, '""');
      return needsQuotes ? `"${escaped}"` : escaped;
    };
    const lines = [header.join(',')];
    for (const r of rows) {
      lines.push([
        r.source,
        r.site,
        r.sourceKind,
        r.text,
        r.current,
        r.suggested,
        r.linkClass,
      ].map(escape).join(','));
    }
    const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `canonicalization-report-${this.state.selectedJobId ?? 'export'}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  /**
   * Walk the loaded results and return every error that occurred —
   * both site-level (the entire site failed to scan) and per-page /
   * per-file (a specific page or doc couldn't be parsed). Errored
   * pages/files don't show up in the link table because they have
   * 0 links, so this is the only way for the user to see them.
   */
  private _collectErrors(): IScanErrorRow[] {
    const out: IScanErrorRow[] = [];
    const results = this.state.results;
    if (!results) return out;
    for (const site of results.sites) {
      if (site.error) {
        out.push({
          site: site.site,
          kind: 'site',
          target: site.site,
          error: site.error,
        });
      }
      for (const page of site.pages) {
        if (page.parseError) {
          out.push({
            site: site.site,
            kind: 'page',
            target: page.pageTitle || page.pageUrl,
            url: page.pageUrl,
            error: page.parseError,
          });
        }
      }
    }
    return out;
  }

  /**
   * Flatten the per-site / per-page results into one row per link, then
   * apply the active filters. Returned rows are stable-ordered by
   * (site, page, link index) so the UI doesn't jump around. Only the
   * primary (selected) job is included — companion rows live exclusively
   * in the canonical index for sibling display.
   */
  private _filteredRows(): ILinkRow[] {
    const { results, selectedJobId, selectedJob, filterClass, filterSite, filterSource, searchText } = this.state;
    if (!results || !selectedJobId) return [];
    const fallbackKind: 'pages' | 'documents' = selectedJob?.kind ?? 'pages';
    const search = searchText.trim().toLowerCase();
    const out: ILinkRow[] = [];
    for (const site of results.sites) {
      if (filterSite && site.site !== filterSite) continue;
      for (const page of site.pages) {
        const pageKind: 'pages' | 'documents' = page.sourceKind ?? fallbackKind;
        for (let i = 0; i < page.links.length; i++) {
          const link = page.links[i];
          if (filterClass && link.linkClass !== filterClass) continue;
          if (filterSource && link.source !== filterSource) continue;
          if (search) {
            const hay = `${page.pageTitle}\n${link.url}\n${link.text ?? ''}`.toLowerCase();
            if (hay.indexOf(search) === -1) continue;
          }
          out.push(this._makeRow(site.site, page, link, i, selectedJobId, pageKind));
        }
      }
    }
    return out;
  }

  private _exportCsv = (): void => {
    const rows = this._filteredRows();
    const header = ['Site', 'Page', 'PageUrl', 'LinkText', 'LinkUrl', 'LinkClass', 'LinkSource', 'WebPartInstanceId', 'NormalizedKey'];
    const escape = (s: string): string => {
      if (s == null) return '';
      const needsQuotes = /[",\r\n]/.test(s);
      const escaped = s.replace(/"/g, '""');
      return needsQuotes ? `"${escaped}"` : escaped;
    };
    const lines = [header.join(',')];
    for (const r of rows) {
      lines.push([
        r.site,
        r.page.pageTitle,
        r.page.pageUrl,
        r.link.text ?? '',
        r.link.url,
        r.link.linkClass,
        r.link.source,
        r.link.webPartInstanceId ?? '',
        r.link.normalizedKey,
      ].map(escape).join(','));
    }
    const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `link-inventory-${this.state.selectedJobId ?? 'export'}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // --- Backlinks column management ---

  private _rebuildBacklinksIndex = async (): Promise<void> => {
    this.setState({ rebuildingBacklinks: true, rebuildError: '', rebuildResult: undefined });
    try {
      const res = await this.props.service.rebuildBacklinksIndex();
      this.setState({ rebuildResult: res });
    } catch (e) {
      this.setState({ rebuildError: (e as Error).message });
    } finally {
      this.setState({ rebuildingBacklinks: false });
    }
  };

  private _exportBacklinksCsv = async (opts?: { siteUrl?: string; libraryTitle?: string }): Promise<void> => {
    this.setState({ exportingBacklinks: true });
    try {
      await this.props.service.downloadBacklinksCsv(opts);
    } catch (e) {
      // Surface via the dialog error when a per-library export is in
      // flight; otherwise use the generic error state.
      if (opts?.siteUrl) {
        this.setState({ backlinksError: (e as Error).message });
      } else {
        this.setState({ rebuildError: (e as Error).message });
      }
    } finally {
      this.setState({ exportingBacklinks: false });
    }
  };

  private _enableBacklinksColumn = async (): Promise<void> => {
    const { backlinksSiteUrl, backlinksLibraryTitle } = this.state;
    if (!backlinksSiteUrl || !backlinksLibraryTitle.trim()) {
      this.setState({ backlinksError: 'Site and library title are required' });
      return;
    }
    this.setState({ backlinksBusy: true, backlinksError: '', backlinksEnableResult: undefined });
    try {
      const res = await this.props.service.enableBacklinksColumn(backlinksSiteUrl, backlinksLibraryTitle.trim());
      this.setState({ backlinksEnableResult: res });
    } catch (e) {
      this.setState({ backlinksError: (e as Error).message });
    } finally {
      this.setState({ backlinksBusy: false });
    }
  };

  private _renderBacklinksDialog(): React.ReactElement | null {
    const {
      backlinksDialogOpen,
      backlinksSiteUrl,
      backlinksLibraryTitle,
      backlinksBusy,
      backlinksError,
      backlinksEnableResult,
      exportingBacklinks,
      sitePickerSites,
    } = this.state;
    if (!backlinksDialogOpen) return null;

    const siteOptions: IDropdownOption[] = [
      { key: '', text: 'Select a site...' },
      ...sitePickerSites
        .slice()
        .sort((a, b) => a.title.localeCompare(b.title))
        .map((s) => ({ key: s.absoluteUrl, text: `${s.title} (${s.serverRelativeUrl})` })),
    ];

    return (
      <Dialog
        hidden={!backlinksDialogOpen}
        onDismiss={() => !backlinksBusy && !exportingBacklinks && this.setState({ backlinksDialogOpen: false })}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Backlinks column',
          subText:
            'Create the Backlinks column on a document library. The column shows, per file, ' +
            'which pages and documents link to that file — gated per user. Non-admins see ' +
            'only source sites they can read.',
        }}
        modalProps={{ isBlocking: true, styles: { main: { maxWidth: 640 } } }}
      >
        <Stack tokens={{ childrenGap: 12 }}>
          <Dropdown
            label="Site"
            options={siteOptions}
            selectedKey={backlinksSiteUrl}
            onChange={(_, opt) => this.setState({ backlinksSiteUrl: (opt?.key as string) || '', backlinksError: '' })}
            disabled={backlinksBusy || exportingBacklinks}
          />
          <TextField
            label="Library title"
            value={backlinksLibraryTitle}
            onChange={(_, v) => this.setState({ backlinksLibraryTitle: v ?? '', backlinksError: '' })}
            placeholder="Shared Documents"
            disabled={backlinksBusy || exportingBacklinks}
            description="Exact library title as it appears in SharePoint (case-sensitive)."
          />
          {backlinksError && (
            <MessageBar messageBarType={MessageBarType.error}>{backlinksError}</MessageBar>
          )}
          {backlinksEnableResult && (
            <MessageBar messageBarType={MessageBarType.success}>
              Column "{backlinksEnableResult.columnDisplayName}" created on {backlinksEnableResult.libraryTitle}.
              {backlinksEnableResult.customizerRegistered
                ? ' Customizer registered.'
                : ' Customizer not yet wired (BACKLINK_CUSTOMIZER_COMPONENT_ID missing).'}
              {backlinksEnableResult.addedToDefaultView
                ? ' Added to default view.'
                : ' You may need to add it to the view manually.'}
            </MessageBar>
          )}
        </Stack>
        <DialogFooter>
          <PrimaryButton
            text={backlinksBusy ? 'Enabling...' : 'Enable column'}
            onClick={() => { void this._enableBacklinksColumn(); }}
            disabled={backlinksBusy || exportingBacklinks || !backlinksSiteUrl || !backlinksLibraryTitle.trim()}
          />
          <DefaultButton
            text={exportingBacklinks ? 'Exporting...' : 'Export CSV for this library'}
            onClick={() => {
              void this._exportBacklinksCsv({
                siteUrl: backlinksSiteUrl,
                libraryTitle: backlinksLibraryTitle.trim(),
              });
            }}
            disabled={backlinksBusy || exportingBacklinks || !backlinksSiteUrl || !backlinksLibraryTitle.trim()}
          />
          <DefaultButton
            text="Close"
            onClick={() => this.setState({ backlinksDialogOpen: false })}
            disabled={backlinksBusy || exportingBacklinks}
          />
        </DialogFooter>
      </Dialog>
    );
  }

  private _renderBacklinksRebuildStatus(): React.ReactElement | null {
    const { rebuildResult, rebuildError } = this.state;
    if (!rebuildResult && !rebuildError) return null;
    if (rebuildError) {
      return (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => this.setState({ rebuildError: '' })}
        >
          Backlinks index rebuild failed: {rebuildError}
        </MessageBar>
      );
    }
    const r = rebuildResult!;
    return (
      <MessageBar
        messageBarType={MessageBarType.success}
        onDismiss={() => this.setState({ rebuildResult: undefined })}
      >
        Backlinks index rebuilt — {r.sitesTotal} sites, {r.sitesWithPageData} with page data,
        {' '}{r.sitesWithDocData} with doc data. Built at {new Date(r.builtAt).toLocaleString()}.
      </MessageBar>
    );
  }

  public render(): React.ReactElement {
    const { loading, isAdmin, jobs, selectedJob, results, resultsLoading, triggering, polling } = this.state;

    if (loading) return <Spinner size={SpinnerSize.large} label="Loading link inventory..." />;

    // Doc scans are read-only — the writer pokes Site Pages list items
    // via REST and has no path for rewriting hyperlinks inside binary
    // OOXML/PDF files. Hide all write-back affordances accordingly.
    const isDocsJob = selectedJob?.kind === 'documents';
    const isWritable = !isDocsJob;

    const jobOptions: IDropdownOption[] = (() => {
      const byId = new Map(jobs.map((j) => [j.jobId, j]));
      const consumedAsSibling = new Set<string>();
      const opts: IDropdownOption[] = [];
      const formatCounts = (j: ILinkInventoryJob): string =>
        j.kind === 'documents'
          ? (j.previewOnly
            ? `${j.filesTotal ?? 0} files enumerated`
            : `${j.filesCompleted ?? 0}/${j.filesTotal ?? 0}f, ${j.linksTotal}L`)
          : `${j.pagesTotal}p, ${j.linksTotal}L`;
      for (const j of jobs) {
        if (consumedAsSibling.has(j.jobId)) continue;
        const sib = j.siblingJobId ? byId.get(j.siblingJobId) : undefined;
        const startedLocal = new Date(j.startedAt).toLocaleString();
        if (sib) {
          // Unified pair — render one entry. Always key on the page job
          // so selection lands on the writable side.
          consumedAsSibling.add(sib.jobId);
          const pages = j.kind === 'pages' ? j : sib;
          const docs = j.kind === 'documents' ? j : sib;
          const label =
            `${startedLocal} — pages: ${pages.status} (${formatCounts(pages)}) · ` +
            `docs: ${docs.status} (${formatCounts(docs)})`;
          opts.push({ key: pages.jobId, text: label });
        } else {
          let tag: string;
          if (j.kind === 'documents') {
            tag = j.previewOnly ? '[DOCS PREVIEW]' : '[DOCS]';
          } else {
            tag = '[PAGES]';
          }
          const label = `${tag} ${startedLocal} — ${j.status} (${formatCounts(j)})`;
          opts.push({ key: j.jobId, text: label });
        }
      }
      return opts;
    })();

    const siteOptions: IDropdownOption[] = [
      { key: '', text: 'All sites' },
      ...(results?.sites.map((s) => ({ key: s.site, text: s.site })) ?? []),
    ];

    const sourceSet = new Set<string>();
    if (results) for (const s of results.sites) for (const p of s.pages) for (const l of p.links) sourceSet.add(l.source);
    const sourceList: string[] = [];
    sourceSet.forEach((s) => sourceList.push(s));
    sourceList.sort();
    const sourceOptions: IDropdownOption[] = [
      { key: '', text: 'All sources' },
      ...sourceList.map((s) => ({ key: s, text: s })),
    ];

    const rows = this._filteredRows();

    const commandItems: ICommandBarItemProps[] = [];
    if (isAdmin) {
      commandItems.push({
        key: 'scan-unified',
        text: triggering ? 'Triggering...' : 'Run scan',
        iconProps: { iconName: 'PlayResume' },
        disabled: triggering || polling,
        title: 'Run a paired tenant-wide page + document scan with default settings',
        onClick: () => { void this._triggerUnifiedScan(); },
      });
      commandItems.push({
        key: 'scan-advanced',
        text: 'Advanced scan',
        iconProps: { iconName: 'Settings' },
        disabled: triggering || polling,
        title: 'Page-only or document-only scans with custom site/options',
        subMenuProps: {
          items: [
            {
              key: 'scan',
              text: 'Page-only scan...',
              iconProps: { iconName: 'Play' },
              onClick: () => {
                this.setState({ pageScanOpen: true });
                void this._loadSitePicker();
              },
            },
            {
              key: 'scan-docs',
              text: 'Document-only scan...',
              iconProps: { iconName: 'Documentation' },
              onClick: () => {
                this.setState({ docScanOpen: true });
                void this._loadSitePicker();
              },
            },
          ],
        },
      });
    }
    commandItems.push({
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => { void this._refresh(); },
    });
    commandItems.push({
      key: 'export',
      text: 'Export CSV',
      iconProps: { iconName: 'Download' },
      disabled: rows.length === 0,
      onClick: () => this._exportCsv(),
    });
    commandItems.push({
      key: 'canon-report',
      text: 'Canonicalization report',
      iconProps: { iconName: 'ReportDocument' },
      disabled: !results,
      onClick: () => this.setState({ canonicalReportOpen: true }),
    });
    // Schedule is a global config — admin-only but not gated by the
    // currently-selected job's kind (isWritable depends on !isDocsJob,
    // which is irrelevant here).
    if (isAdmin) {
      commandItems.push({
        key: 'schedule',
        text: 'Schedule',
        iconProps: { iconName: 'Clock' },
        title: 'Configure the daily delta scan',
        onClick: () => { void this._openSchedulePanel(); },
      });
    }
    if (isWritable) {
      commandItems.push({
        key: 'find',
        text: 'Find & Replace',
        iconProps: { iconName: 'FindAndReplace' },
        disabled: !results,
        onClick: () => this.setState({
          findReplaceOpen: true,
          frMode: 'substring',
          frModeLabel: '',
          frResponse: undefined,
          frApplied: false,
          frError: '',
        }),
      });
    }
    if (isAdmin) {
      commandItems.push({
        key: 'backlinks',
        text: 'Backlinks column',
        iconProps: { iconName: 'Link' },
        subMenuProps: {
          items: [
            {
              key: 'enable',
              text: 'Enable on library...',
              iconProps: { iconName: 'AddTo' },
              onClick: () => {
                this.setState({
                  backlinksDialogOpen: true,
                  backlinksError: '',
                  backlinksEnableResult: undefined,
                });
                void this._loadSitePicker();
              },
            },
            {
              key: 'rebuild',
              text: this.state.rebuildingBacklinks ? 'Rebuilding...' : 'Rebuild index',
              iconProps: { iconName: 'Refresh' },
              disabled: this.state.rebuildingBacklinks,
              onClick: () => { void this._rebuildBacklinksIndex(); },
            },
            {
              key: 'export-tenant',
              text: this.state.exportingBacklinks ? 'Exporting...' : 'Export CSV (tenant-wide)',
              iconProps: { iconName: 'Download' },
              disabled: this.state.exportingBacklinks,
              onClick: () => { void this._exportBacklinksCsv(); },
            },
          ],
        },
      });
      commandItems.push({
        key: 'delete',
        text: this.state.deleting ? 'Deleting...' : 'Delete this scan',
        iconProps: { iconName: 'Delete' },
        disabled: !selectedJob || this.state.deleting || polling,
        onClick: () => this.setState({ deleteConfirmOpen: true }),
      });
    }

    const columns: IColumn[] = [];
    columns.push({
      key: 'class', name: 'Class', fieldName: 'linkClass', minWidth: 90, maxWidth: 130, isResizable: true,
      onRender: (item: ILinkRow) => <span style={this._classBadgeStyle(item.link.linkClass)}>{item.link.linkClass}</span>,
    });
    // Source column is page-specific (text web part / layout web part / etc).
    // Doc extractor sources (ooxml-rels / pdf-annot) are uninteresting to admins.
    if (!isDocsJob) {
      columns.push({
        key: 'source', name: 'Source', fieldName: 'source', minWidth: 80, maxWidth: 110, isResizable: true,
        onRender: (item: ILinkRow) => <Text variant="small">{item.link.source}</Text>,
      });
    }
    columns.push({
      key: 'site', name: 'Site', fieldName: 'site', minWidth: 130, maxWidth: 220, isResizable: true,
      onRender: (item: ILinkRow) => <Text variant="small">{item.site}</Text>,
    });
    columns.push({
      key: 'page',
      name: isDocsJob ? 'File' : 'Page',
      fieldName: 'pageTitle',
      minWidth: 160,
      maxWidth: 280,
      isResizable: true,
      onRender: (item: ILinkRow) => {
        if (isDocsJob) {
          // Render as an icon + click-out link to open the file in SPO.
          // pageUrl is server-relative for doc-scan rows (it's the file's
          // serverRelativeUrl), so prefix with TENANT_ORIGIN.
          const href = item.page.pageUrl.startsWith('http')
            ? item.page.pageUrl
            : `${TENANT_ORIGIN}${item.page.pageUrl}`;
          return (
            <span style={{ display: 'inline-flex', alignItems: 'center', minWidth: 0 }}>
              {this._renderFileIcon(item.page.pageTitle, 16)}
              <a
                href={href}
                target="_blank"
                rel="noopener noreferrer"
                title={item.page.pageUrl}
                style={{
                  color: 'inherit',
                  textDecoration: 'underline',
                  fontSize: 12,
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap',
                }}
                onClick={(e) => e.stopPropagation()}
              >
                {item.page.pageTitle}
              </a>
            </span>
          );
        }
        return <Text variant="small" title={item.page.pageUrl}>{item.page.pageTitle}</Text>;
      },
    });
    columns.push({
      key: 'text', name: 'Link text', minWidth: 120, maxWidth: 220, isResizable: true,
      onRender: (item: ILinkRow) => (
        <Text variant="small" title={item.link.text}>{item.link.text ?? <em style={{ opacity: 0.5 }}>(none)</em>}</Text>
      ),
    });
    columns.push({
      key: 'url', name: 'Target URL', fieldName: 'url', minWidth: 220, isResizable: true,
      onRender: (item: ILinkRow) => {
        const url = item.link.url;
        const display = url.length > 80 ? url.slice(0, 78) + '…' : url;
        // Only http(s) URLs are click-safe. mailto:, tel:, javascript:,
        // anchor-only, and relative paths render as plain text so we
        // don't trigger mail clients or security warnings on a click.
        const isClickable = /^https?:\/\//i.test(url);
        if (!isClickable) {
          return (
            <span style={{ fontFamily: 'Consolas, monospace', fontSize: 11 }} title={url}>
              {display}
            </span>
          );
        }
        return (
          <a
            href={url}
            target="_blank"
            rel="noopener noreferrer"
            title={url}
            style={{
              fontFamily: 'Consolas, monospace',
              fontSize: 11,
              color: 'inherit',
              textDecoration: 'underline',
            }}
            // Stop the row's onActiveItemChanged from also firing —
            // we want a click on the link to navigate, not open the
            // detail panel underneath.
            onClick={(e) => e.stopPropagation()}
          >
            {display}
          </a>
        );
      },
    });

    return (
      <div style={{ padding: '10px 0' }}>
        <Stack tokens={{ childrenGap: 12 }}>
          <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="end">
            <Dropdown
              label="Scan"
              options={jobOptions}
              selectedKey={this.state.selectedJobId}
              onChange={this._onJobChange}
              styles={{ root: { width: 460 } }}
              disabled={jobs.length === 0}
              placeholder={jobs.length === 0 ? 'No scans yet' : undefined}
            />
            {selectedJob && (
              <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="small">
                  {selectedJob.kind === 'documents' ? 'Document scan' : 'Page scan'} · Status: <strong>{selectedJob.status}</strong>
                  {selectedJob.resultsAvailable && results && (
                    <> · {results.userScopedTotals.sites}/{selectedJob.sitesTotal} sites visible to you</>
                  )}
                </Text>
                <Text variant="small">
                  {selectedJob.kind === 'documents'
                    ? `${selectedJob.filesCompleted ?? 0}/${selectedJob.filesTotal ?? 0} files · ${selectedJob.linksTotal} links`
                    : `${selectedJob.pagesTotal} pages · ${selectedJob.linksTotal} links`}
                  {selectedJob.errorCount > 0 && ` · ${selectedJob.errorCount} errors`}
                </Text>
              </Stack>
            )}
          </Stack>

          {selectedJob && (selectedJob.status === 'queued' || selectedJob.status === 'running') && (
            (() => {
              const isDocs = selectedJob.kind === 'documents';
              const filesTotal = selectedJob.filesTotal ?? 0;
              const filesDone = selectedJob.filesCompleted ?? 0;
              // For doc scans, fall back to site progress while filesTotal=0 (enumerate phase)
              const inEnumeratePhase = isDocs && filesTotal === 0;
              const label = isDocs
                ? `Document scan: ${selectedJob.currentSite ?? '...'}`
                : `Scanning ${selectedJob.currentSite ?? '...'}`;
              const description = isDocs
                ? inEnumeratePhase
                  ? `Discovering files — ${selectedJob.sitesCompleted}/${selectedJob.sitesTotal} sites enumerated`
                  : `${filesDone}/${filesTotal} files · ${selectedJob.linksTotal} links so far`
                : `${selectedJob.sitesCompleted} of ${selectedJob.sitesTotal} sites · ${selectedJob.pagesTotal} pages so far`;
              const pct = isDocs
                ? inEnumeratePhase
                  ? (selectedJob.sitesTotal > 0 ? selectedJob.sitesCompleted / selectedJob.sitesTotal : 0)
                  : (filesTotal > 0 ? filesDone / filesTotal : 0)
                : (selectedJob.sitesTotal > 0 ? selectedJob.sitesCompleted / selectedJob.sitesTotal : 0);
              return (
                <ProgressIndicator
                  label={label}
                  description={description}
                  percentComplete={pct}
                />
              );
            })()
          )}

          {results && results.droppedSites.length > 0 && (
            <MessageBar messageBarType={MessageBarType.info}>
              <Stack tokens={{ childrenGap: 4 }}>
                <span>
                  <strong>{results.droppedSites.length}</strong> site(s) hidden —
                  you don't have read access.{' '}
                  <a
                    href="#"
                    onClick={(e) => {
                      e.preventDefault();
                      this.setState({ droppedSitesExpanded: !this.state.droppedSitesExpanded });
                    }}
                    style={{ textDecoration: 'underline', color: 'inherit', cursor: 'pointer' }}
                  >
                    {this.state.droppedSitesExpanded ? 'Hide list' : 'Show list'}
                  </a>
                </span>
                {this.state.droppedSitesExpanded && (
                  <div
                    style={{
                      maxHeight: 200,
                      overflowY: 'auto',
                      fontFamily: "'Consolas', 'Courier New', monospace",
                      fontSize: 11,
                      padding: '6px 8px',
                      background: 'rgba(127,127,127,0.12)',
                      color: 'inherit',
                      border: '1px solid rgba(127,127,127,0.4)',
                      borderRadius: 2,
                      lineHeight: 1.6,
                    }}
                  >
                    {results.droppedSites.slice().sort().join(' · ')}
                  </div>
                )}
              </Stack>
            </MessageBar>
          )}

          {!isAdmin && jobs.length === 0 && (
            <MessageBar messageBarType={MessageBarType.info}>
              No scans have been run yet. Ask a SharePoint Redirect Manager Admin to trigger one.
            </MessageBar>
          )}

          {isDocsJob && results && !selectedJob?.previewOnly && (
            <MessageBar messageBarType={MessageBarType.info}>
              <strong>Read-only.</strong> Document scans expose links inside Word, Excel,
              PowerPoint, and PDF files. The Find &amp; Replace and "Fix this link" actions
              don't apply because we can't rewrite hyperlinks inside binary files —
              click a file name to open it in SPO and fix the link manually.
            </MessageBar>
          )}

          <CommandBar items={commandItems} />
          {this._renderBacklinksRebuildStatus()}

          {this.state.previewResults && this._renderPreviewView(this.state.previewResults)}

          {results && (
            <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
              <Dropdown
                label="Class"
                options={LINK_CLASS_OPTIONS}
                selectedKey={this.state.filterClass}
                onChange={(_e, opt) => this.setState({ filterClass: opt ? String(opt.key) : '' })}
                styles={{ root: { width: 200 } }}
              />
              <Dropdown
                label="Site"
                options={siteOptions}
                selectedKey={this.state.filterSite}
                onChange={(_e, opt) => this.setState({ filterSite: opt ? String(opt.key) : '' })}
                styles={{ root: { width: 240 } }}
              />
              <Dropdown
                label="Source"
                options={sourceOptions}
                selectedKey={this.state.filterSource}
                onChange={(_e, opt) => this.setState({ filterSource: opt ? String(opt.key) : '' })}
                styles={{ root: { width: 200 } }}
              />
              <Stack.Item grow>
                <SearchBox
                  placeholder="Search by URL, page, or link text..."
                  value={this.state.searchText}
                  onChange={(_e, val) => this.setState({ searchText: val ?? '' })}
                  styles={{ root: { marginTop: 28 } }}
                />
              </Stack.Item>
            </Stack>
          )}

          {resultsLoading && <Spinner size={SpinnerSize.medium} label="Loading results..." />}

          {results && !resultsLoading && this._renderErrorsPanel()}

          {results && !resultsLoading && (
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="small">{rows.length.toLocaleString()} link(s) shown</Text>
              <DetailsList
                key={this.state.detailsListKey}
                items={rows}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                onActiveItemChanged={(item) => this.setState({ detailRow: item as ILinkRow })}
                compact
              />
            </Stack>
          )}
        </Stack>

        <Panel
          isOpen={!!this.state.detailRow}
          onDismiss={() => this.setState({
            detailRow: undefined,
            detailsListKey: this.state.detailsListKey + 1,
          })}
          headerText="Link details"
          type={PanelType.medium}
        >
          {this.state.detailRow && this._renderDetailPanel(this.state.detailRow)}
        </Panel>

        <Panel
          isOpen={this.state.findReplaceOpen}
          onDismiss={() => this.setState({ findReplaceOpen: false })}
          headerText="Find & Replace"
          type={PanelType.medium}
        >
          {this._renderFindReplacePanel()}
        </Panel>

        <Panel
          isOpen={this.state.docScanOpen}
          onDismiss={() => this.setState({ docScanOpen: false })}
          headerText="Run document scan"
          type={PanelType.medium}
        >
          {this._renderDocScanPanel()}
        </Panel>

        <Panel
          isOpen={this.state.pageScanOpen}
          onDismiss={() => this.setState({ pageScanOpen: false })}
          headerText="Run page scan"
          type={PanelType.medium}
        >
          {this._renderPageScanPanel()}
        </Panel>

        <Panel
          isOpen={this.state.canonicalReportOpen}
          onDismiss={() => this.setState({ canonicalReportOpen: false })}
          headerText="Canonicalization report"
          type={PanelType.large}
        >
          {this._renderCanonicalReportPanel()}
        </Panel>

        <Panel
          isOpen={this.state.scheduleOpen}
          onDismiss={() => this.setState({ scheduleOpen: false })}
          headerText="Daily scan schedule"
          type={PanelType.medium}
        >
          {this._renderSchedulePanel()}
        </Panel>

        <Dialog
          hidden={!this.state.deleteConfirmOpen}
          onDismiss={() => !this.state.deleting && this.setState({ deleteConfirmOpen: false })}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Delete this scan?',
            subText: selectedJob
              ? `This will permanently delete the ${selectedJob.kind === 'documents' ? 'document' : 'page'} ` +
                `scan from ${new Date(selectedJob.startedAt).toLocaleString()}, including all results, ` +
                `partial blobs, and manifests. This cannot be undone.`
              : 'This will permanently delete the selected scan. This cannot be undone.',
          }}
          modalProps={{ isBlocking: true }}
        >
          <DialogFooter>
            <PrimaryButton
              text={this.state.deleting ? 'Deleting…' : 'Delete'}
              onClick={() => { void this._deleteSelectedJob(); }}
              disabled={this.state.deleting}
              styles={{ root: { background: '#a4262c', borderColor: '#a4262c' }, rootHovered: { background: '#8f1f25', borderColor: '#8f1f25' } }}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ deleteConfirmOpen: false })}
              disabled={this.state.deleting}
            />
          </DialogFooter>
        </Dialog>

        {this._renderBacklinksDialog()}
      </div>
    );
  }

  /**
   * Delete the currently-selected scan job and its blobs. Selects
   * the next-most-recent job after deletion (or clears the panel if
   * none remain).
   */
  private _deleteSelectedJob = async (): Promise<void> => {
    const { selectedJobId, jobs } = this.state;
    if (!selectedJobId) return;
    this.setState({ deleting: true });
    try {
      await this.props.service.deleteJob(selectedJobId);
      // Refresh the job list and pick the next one (or clear if empty).
      const fresh = await this.props.service.listJobs();
      const next = fresh.find((j) => j.resultsAvailable) ?? fresh[0];
      this.setState({
        deleting: false,
        deleteConfirmOpen: false,
        jobs: fresh,
        selectedJobId: next?.jobId,
        selectedJob: next,
        results: undefined,
        previewResults: undefined,
        previewSelectedSites: new Set(),
        canonicalIndex: new Map(),
      });
      if (next?.resultsAvailable) {
        await this._loadResults(next.jobId);
      } else if (next && (next.status === 'queued' || next.status === 'running')) {
        this._startPolling(next.jobId);
      }
    } catch (err) {
      this.setState({ deleting: false });
      this.props.onError(`Delete scan: ${(err as Error).message}`);
    }
    // Suppress unused-warning for jobs (we use it via state above).
    void jobs;
  };

  /**
   * File-type icon for a given filename, using the Microsoft Fluent
   * file type icon CDN — the same SVGs Modern SPO renders for OneDrive
   * and library tile views. No new package dependency needed.
   *
   * Sizes available: 16, 20, 24, 32, 40, 48, 64, 96.
   *
   * Falls back to `genericfile` when the extension isn't recognized.
   */
  private _fileTypeIconUrl(fileName: string, size: 16 | 20 | 24 | 32 = 16): string {
    const m = /\.([a-z0-9]+)$/i.exec(fileName);
    const ext = (m?.[1] ?? 'genericfile').toLowerCase();
    return `https://res.cdn.office.net/files/fabric-cdn-prod_20221209.001/assets/item-types/${size}/${ext}.svg`;
  }

  private _renderFileIcon(fileName: string, size: 16 | 20 | 24 | 32 = 16): React.ReactElement {
    return (
      <img
        src={this._fileTypeIconUrl(fileName, size)}
        alt=""
        width={size}
        height={size}
        style={{ verticalAlign: 'middle', marginRight: 6, flexShrink: 0 }}
        // If the CDN doesn't have this extension, fall back to the
        // generic file icon so we never show a broken-image glyph.
        onError={(e) => {
          const img = e.currentTarget;
          if (!img.dataset.fallback) {
            img.dataset.fallback = '1';
            img.src = `https://res.cdn.office.net/files/fabric-cdn-prod_20221209.001/assets/item-types/${size}/genericfile.svg`;
          }
        }}
      />
    );
  }

  /**
   * Theme-neutral input style. The webpart renders inside SP pages
   * which can be light OR dark themed; using semi-transparent grey
   * overlays + inherited text color works in both.
   */
  private _inputStyle(disabled: boolean): React.CSSProperties {
    return {
      width: '100%',
      padding: 6,
      fontFamily: 'Consolas, monospace',
      fontSize: 12,
      background: 'rgba(127,127,127,0.12)',
      color: 'inherit',
      border: '1px solid rgba(127,127,127,0.4)',
      borderRadius: 2,
      opacity: disabled ? 0.6 : 1,
    };
  }

  private _textareaStyle(minHeight: number, disabled?: boolean): React.CSSProperties {
    return {
      width: '100%',
      minHeight,
      padding: 6,
      fontFamily: 'Consolas, monospace',
      fontSize: 12,
      background: 'rgba(127,127,127,0.12)',
      color: 'inherit',
      border: '1px solid rgba(127,127,127,0.4)',
      borderRadius: 2,
      resize: 'vertical',
      opacity: disabled ? 0.5 : 1,
    };
  }

  private _classBadgeStyle(linkClass: string): React.CSSProperties {
    // Use semi-transparent foreground tint as background — works in
    // both light and dark themes because the contrast is generated
    // from a single foreground color rather than fixed light bg + dark
    // text (which collapses on dark theme).
    const fg: Record<string, string> = {
      'malformed-spo-link': '#d13438',
      'onprem':        '#e0494f',
      'sharing-link':  '#d9a32a',
      'doc-aspx':      '#d9a32a',
      'office-online': '#3aaa3a',
      'spo-internal':  '#3a96dd',
      'relative':      '#5fa9e8',
      'external':      '#9a9a9a',
      'mailto':        '#9a9a9a',
      'tel':           '#9a9a9a',
      'anchor-only':   '#7a7a7a',
      'javascript':    '#e0494f',
      'unknown':       '#9a9a9a',
    };
    const color = fg[linkClass] ?? fg.unknown;
    return {
      background: this._tintBg(color),
      color,
      padding: '2px 8px',
      borderRadius: 10,
      fontSize: 11,
      fontWeight: 600,
      border: `1px solid ${this._tintBorder(color)}`,
    };
  }

  /** Convert a hex color to a low-opacity semi-transparent rgba string. */
  private _tintBg(hex: string): string {
    const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex);
    if (!m) return 'rgba(127,127,127,0.15)';
    return `rgba(${parseInt(m[1], 16)},${parseInt(m[2], 16)},${parseInt(m[3], 16)},0.18)`;
  }

  private _tintBorder(hex: string): string {
    const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex);
    if (!m) return 'rgba(127,127,127,0.4)';
    return `rgba(${parseInt(m[1], 16)},${parseInt(m[2], 16)},${parseInt(m[3], 16)},0.45)`;
  }

  /**
   * Collapsible "Errors" panel — shown above the link table when the
   * loaded results contain any failed sites/pages/files. Click the
   * MessageBar header to expand the list.
   */
  private _renderErrorsPanel(): React.ReactElement | null {
    const errors = this._collectErrors();
    if (errors.length === 0) return null;
    const expanded = this.state.errorsExpanded;
    const isDocs = this.state.selectedJob?.kind === 'documents';
    const itemNoun = isDocs ? 'file' : 'page';

    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        <Stack tokens={{ childrenGap: 6 }}>
          <span>
            <strong>{errors.length}</strong> error(s) during scan —{' '}
            <a
              href="#"
              onClick={(e) => {
                e.preventDefault();
                this.setState({ errorsExpanded: !expanded });
              }}
              style={{ textDecoration: 'underline', color: 'inherit', cursor: 'pointer' }}
            >
              {expanded ? 'Hide details' : 'Show details'}
            </a>
          </span>
          {expanded && (
            <div
              style={{
                maxHeight: 320,
                overflowY: 'auto',
                background: 'rgba(127,127,127,0.08)',
                color: 'inherit',
                border: '1px solid rgba(127,127,127,0.4)',
                borderRadius: 2,
                padding: 4,
              }}
            >
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                <thead>
                  <tr style={{ borderBottom: '1px solid rgba(127,127,127,0.4)', textAlign: 'left' }}>
                    <th style={{ padding: '4px 8px', width: 60 }}>Kind</th>
                    <th style={{ padding: '4px 8px' }}>Site</th>
                    <th style={{ padding: '4px 8px' }}>{isDocs ? 'File' : 'Page'}</th>
                    <th style={{ padding: '4px 8px' }}>Error</th>
                  </tr>
                </thead>
                <tbody>
                  {errors.map((e, i) => {
                    const isSiteErr = e.kind === 'site';
                    const fullHref = e.url
                      ? (e.url.startsWith('http') ? e.url : `${TENANT_ORIGIN}${e.url}`)
                      : undefined;
                    return (
                      <tr
                        key={`${e.site}:${e.kind}:${i}`}
                        style={{ borderBottom: '1px solid rgba(127,127,127,0.2)' }}
                      >
                        <td style={{ padding: '4px 8px', verticalAlign: 'top' }}>
                          <span
                            style={{
                              background: isSiteErr ? 'rgba(224,73,79,0.18)' : 'rgba(217,163,42,0.18)',
                              color: isSiteErr ? '#e0494f' : '#d9a32a',
                              border: isSiteErr ? '1px solid rgba(224,73,79,0.45)' : '1px solid rgba(217,163,42,0.45)',
                              padding: '1px 6px',
                              borderRadius: 8,
                              fontSize: 10,
                              fontWeight: 700,
                              letterSpacing: 0.4,
                              textTransform: 'uppercase',
                            }}
                          >
                            {isSiteErr ? 'site' : itemNoun}
                          </span>
                        </td>
                        <td style={{ padding: '4px 8px', verticalAlign: 'top', fontFamily: 'Consolas, monospace', fontSize: 11 }}>
                          {e.site}
                        </td>
                        <td style={{ padding: '4px 8px', verticalAlign: 'top' }}>
                          {isSiteErr ? (
                            <em style={{ opacity: 0.6 }}>(site-level)</em>
                          ) : fullHref ? (
                            <a
                              href={fullHref}
                              target="_blank"
                              rel="noopener noreferrer"
                              style={{ color: 'inherit', textDecoration: 'underline' }}
                              title={e.url}
                            >
                              {e.target}
                            </a>
                          ) : (
                            <span title={e.url}>{e.target}</span>
                          )}
                        </td>
                        <td style={{ padding: '4px 8px', verticalAlign: 'top', fontFamily: 'Consolas, monospace', fontSize: 11, wordBreak: 'break-word' }}>
                          {e.error}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </Stack>
      </MessageBar>
    );
  }

  /**
   * Render a sibling row list. Each row shows a small kind badge
   * (PAGES / DOCS), a file-type icon for doc rows, the link class
   * badge, the page/file title, and the URL. Clicking the row opens
   * its detail panel.
   */
  private _renderSiblingList(siblings: ILinkRow[]): React.ReactElement {
    return (
      <div
        style={{
          maxHeight: 280,
          overflowY: 'auto',
          border: '1px solid rgba(127,127,127,0.4)',
          borderRadius: 2,
          padding: 6,
          background: 'rgba(127,127,127,0.06)',
        }}
      >
        {siblings.map((sib) => {
          const sibIsDoc = sib.jobKind === 'documents';
          return (
            <div
              key={sib.key}
              style={{
                padding: '6px 8px',
                borderBottom: '1px solid rgba(127,127,127,0.25)',
                cursor: 'pointer',
              }}
              onClick={() => this.setState({ detailRow: sib })}
              title="Open this occurrence"
            >
              <div style={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 6, fontSize: 12 }}>
                <span
                  style={{
                    background: sibIsDoc ? 'rgba(74,144,226,0.18)' : 'rgba(160,116,213,0.18)',
                    color: sibIsDoc ? '#4a90e2' : '#a074d5',
                    border: sibIsDoc ? '1px solid rgba(74,144,226,0.45)' : '1px solid rgba(160,116,213,0.45)',
                    padding: '1px 6px',
                    borderRadius: 8,
                    fontSize: 10,
                    fontWeight: 700,
                    letterSpacing: 0.4,
                  }}
                >
                  {sibIsDoc ? 'DOC' : 'PAGE'}
                </span>
                <span style={this._classBadgeStyle(sib.link.linkClass)}>{sib.link.linkClass}</span>
                {sibIsDoc && this._renderFileIcon(sib.page.pageTitle, 16)}
                <strong>{sib.page.pageTitle}</strong>
                <span style={{ opacity: 0.7 }}>· {sib.site}</span>
              </div>
              {sib.link.text && (
                <div style={{ fontSize: 11, opacity: 0.7, marginTop: 2 }}>{sib.link.text}</div>
              )}
              <div style={{ fontFamily: 'Consolas, monospace', fontSize: 11, marginTop: 2 }}>
                {sib.link.url.length > 100 ? sib.link.url.slice(0, 98) + '…' : sib.link.url}
              </div>
            </div>
          );
        })}
      </div>
    );
  }

  private _renderDetailPanel(row: ILinkRow): React.ReactElement {
    const hasSuggestion = !!row.link.suggestion && row.link.suggestion.length > 0;
    const rowIsDoc = row.jobKind === 'documents';
    // Doc rows can't be written back to — buttons hidden.
    const showWriteActions = !rowIsDoc;

    // Find sibling rows: every other row that shares this link's canonical
    // key, across BOTH the primary and companion jobs. The new composite
    // key (jobId:pageUrl:linkIdx) is unique enough that a single
    // `r.key === row.key` comparison correctly identifies the current row.
    const siblings: ILinkRow[] = [];
    if (row.link.canonicalKey) {
      const all = this.state.canonicalIndex.get(row.link.canonicalKey) ?? [];
      for (const r of all) {
        if (r.key === row.key) continue;
        siblings.push(r);
      }
    }

    // Partition siblings by source kind so we can render them in two
    // visually distinct groups: same kind as the current row (page
    // siblings if viewing a page row, doc siblings if viewing a doc
    // row), and the cross-kind partition.
    const sameJobSiblings = siblings.filter((s) => s.jobKind === row.jobKind);
    const otherJobSiblings = siblings.filter((s) => s.jobKind !== row.jobKind);

    // For doc rows, the "page" url is a server-relative file path —
    // make it a click-through to open the file in SPO.
    const sourceHref = row.page.pageUrl.startsWith('http')
      ? row.page.pageUrl
      : `${TENANT_ORIGIN}${row.page.pageUrl}`;

    return (
      <Stack tokens={{ childrenGap: 12 }} style={{ paddingTop: 12 }}>
        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>
            {rowIsDoc ? 'File' : 'Page'}
          </Text>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
            {rowIsDoc && this._renderFileIcon(row.page.pageTitle, 20)}
            {rowIsDoc ? (
              <a
                href={sourceHref}
                target="_blank"
                rel="noopener noreferrer"
                style={{ color: 'inherit', textDecoration: 'underline' }}
              >
                {row.page.pageTitle}
              </a>
            ) : (
              <Text>{row.page.pageTitle}</Text>
            )}
          </Stack>
          <Text variant="small" styles={{ root: { fontFamily: 'Consolas, monospace', opacity: 0.75 } }}>{row.page.pageUrl}</Text>
        </Stack>
        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>Site</Text>
          <Text>{row.site}</Text>
        </Stack>
        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>Link text</Text>
          <Text>{row.link.text ?? <em>(none)</em>}</Text>
        </Stack>
        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>Target URL</Text>
          <textarea
            readOnly
            value={row.link.url}
            style={{
              width: '100%',
              minHeight: 60,
              fontFamily: 'Consolas, monospace',
              fontSize: 12,
              padding: 6,
              background: 'rgba(127,127,127,0.12)',
              color: 'inherit',
              border: '1px solid rgba(127,127,127,0.4)',
              borderRadius: 2,
              resize: 'vertical',
            }}
          />
        </Stack>

        {row.link.linkClass === 'malformed-spo-link' && (
          <Stack tokens={{ childrenGap: 4 }} style={{
            background: 'rgba(209, 52, 56, 0.12)',
            padding: 10,
            borderLeft: '3px solid #d13438',
            borderRadius: 2,
          }}>
            <Text variant="smallPlus" styles={{ root: { color: '#d13438', fontWeight: 600 } }}>
              Malformed SPO link
              {row.link.malformedReason ? ` · ${row.link.malformedReason}` : ''}
            </Text>
            <Text variant="small">
              {row.link.malformedReason
                ? MALFORMED_REASON_LABEL[row.link.malformedReason]
                : 'This link matches a known malformed-URL shape. Review the source page and replace it with a canonical AllItems URL for the correct file.'}
            </Text>
          </Stack>
        )}

        {hasSuggestion && (
          <Stack tokens={{ childrenGap: 4 }} style={{
            background: 'rgba(58, 170, 58, 0.12)',
            padding: 10,
            borderLeft: '3px solid #3aaa3a',
            borderRadius: 2,
          }}>
            <Text variant="smallPlus" styles={{ root: { color: '#3aaa3a', fontWeight: 600 } }}>
              Suggested replacement
            </Text>
            <textarea
              readOnly
              value={row.link.suggestion}
              style={{
              width: '100%',
              minHeight: 60,
              fontFamily: 'Consolas, monospace',
              fontSize: 12,
              padding: 6,
              background: 'rgba(127,127,127,0.12)',
              color: 'inherit',
              border: '1px solid rgba(127,127,127,0.4)',
              borderRadius: 2,
              resize: 'vertical',
            }}
            />
            <Text variant="small" styles={{ root: { opacity: 0.75 } }}>
              This is a stable AllItems URL that won't break under permission changes
              or when the file is moved within its library.
            </Text>
            {showWriteActions ? (
              <button
                onClick={() => this._fixThisLink(row)}
                style={{
                  padding: '8px 24px',
                  background: '#107c10',
                  color: '#fff',
                  border: 'none',
                  borderRadius: 2,
                  fontSize: 14,
                  cursor: 'pointer',
                  marginTop: 6,
                  alignSelf: 'flex-start',
                }}
              >
                Fix this link
              </button>
            ) : (
              <Text variant="small" styles={{ root: { opacity: 0.75, marginTop: 6 } }}>
                <em>Read-only — open the file in SPO to edit the link manually.</em>
              </Text>
            )}
          </Stack>
        )}

        {sameJobSiblings.length > 0 && (
          <Stack tokens={{ childrenGap: 4 }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>
                Also linked from <strong>{sameJobSiblings.length}</strong> other{' '}
                {rowIsDoc ? 'file' : 'page'}(s) in this scan — same target
              </Text>
              {hasSuggestion && showWriteActions && (
                <button
                  onClick={() => this._alignAllForms(row)}
                  style={{
                    padding: '6px 14px',
                    background: '#107c10',
                    color: '#fff',
                    border: 'none',
                    borderRadius: 2,
                    fontSize: 12,
                    cursor: 'pointer',
                  }}
                  title="Replace every form (sharing wrapper, direct, AllItems) with the canonical URL above"
                >
                  Align all to canonical
                </button>
              )}
            </Stack>
            {this._renderSiblingList(sameJobSiblings)}
            <Text variant="small" styles={{ root: { opacity: 0.75 } }}>
              These are different URL forms (sharing wrappers, direct paths, AllItems URLs)
              that all resolve to the same target.
              {showWriteActions && ' Find & Replace will catch them all if you use the suggested replacement above.'}
            </Text>
          </Stack>
        )}

        {otherJobSiblings.length > 0 && (
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>
              Also linked from <strong>{otherJobSiblings.length}</strong>{' '}
              {rowIsDoc ? 'page(s)' : 'document file(s)'} in the latest{' '}
              {rowIsDoc ? 'page' : 'document'} scan — same target
            </Text>
            {this._renderSiblingList(otherJobSiblings)}
            <Text variant="small" styles={{ root: { opacity: 0.75 } }}>
              Cross-scan view — these occurrences come from a different scan kind, so
              they're shown for visibility only and aren't affected by Find &amp; Replace
              from this view.
            </Text>
          </Stack>
        )}
        {row.link.normalizedKey && (
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>Normalized lookup key</Text>
            <Text variant="small" styles={{ root: { fontFamily: 'Consolas, monospace' } }}>
              {row.link.normalizedKey}
            </Text>
          </Stack>
        )}
        <Stack horizontal tokens={{ childrenGap: 16 }}>
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>Link class</Text>
            <span style={this._classBadgeStyle(row.link.linkClass)}>{row.link.linkClass}</span>
          </Stack>
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>Source web part</Text>
            <Text>{row.link.source}</Text>
          </Stack>
        </Stack>
        {row.link.webPartInstanceId && (
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>Web part instance id</Text>
            <Text variant="small" styles={{ root: { fontFamily: 'Consolas, monospace' } }}>{row.link.webPartInstanceId}</Text>
          </Stack>
        )}
        {row.page.etag && (
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>Page ETag (at scan time)</Text>
            <Text variant="small" styles={{ root: { fontFamily: 'Consolas, monospace' } }}>{row.page.etag}</Text>
          </Stack>
        )}
      </Stack>
    );
  }

  /**
   * Open the find/replace panel pre-filled to fix a single link.
   * The find pattern is the link's exact decoded URL, the replace
   * pattern is the suggested canonical form. The user clicks Preview
   * to validate then Apply to write back.
   */
  private _fixThisLink = (row: ILinkRow): void => {
    if (!row.link.suggestion) return;
    this.setState({
      detailRow: undefined,
      detailsListKey: this.state.detailsListKey + 1,
      findReplaceOpen: true,
      frFind: row.link.url,
      frReplace: row.link.suggestion,
      frMode: 'substring',
      frModeLabel: '',
      frResponse: undefined,
      frApplied: false,
      frError: '',
    });
  };

  /**
   * Open the find/replace panel in canonical mode — replaces every
   * URL form (sharing wrapper, direct, AllItems) of the same target
   * with the canonical suggestion. The matching is by canonical key,
   * not substring.
   */
  private _alignAllForms = (row: ILinkRow): void => {
    if (!row.link.suggestion || !row.link.canonicalKey) return;
    this.setState({
      detailRow: undefined,
      detailsListKey: this.state.detailsListKey + 1,
      findReplaceOpen: true,
      frFind: row.link.canonicalKey,
      frReplace: row.link.suggestion,
      frMode: 'canonical',
      frModeLabel: 'Align all forms to canonical',
      frResponse: undefined,
      frApplied: false,
      frError: '',
    });
  };

  private _runReplace = async (dryRun: boolean): Promise<void> => {
    const { frFind, frReplace, frMode, selectedJobId } = this.state;
    if (!selectedJobId || !frFind.trim()) return;
    this.setState({ frBusy: true, frError: '', frResponse: dryRun ? undefined : this.state.frResponse });
    try {
      const response = await this.props.service.replace(selectedJobId, frFind, frReplace, { dryRun, mode: frMode });
      this.setState({
        frBusy: false,
        frResponse: response,
        frApplied: !dryRun && response.summary.applied > 0,
      });
      // After a real-run with any successful applies, refresh the underlying
      // results so the next dry-run reflects the new state.
      if (!dryRun && response.summary.applied > 0) {
        await this._loadResults(selectedJobId);
      }
    } catch (err) {
      this.setState({ frBusy: false, frError: (err as Error).message });
    }
  };

  /**
   * Canonicalization Report panel — table of every non-canonical
   * SPO document link from the loaded job(s), with the four columns
   * the user asked for: file/page, link text, current link,
   * suggested canonical (viewer) link. Includes a CSV export.
   *
   * Spans both the primary and companion (cross-job) results so the
   * report covers links from BOTH page and document files in one
   * shot, regardless of which job is currently selected.
   */
  private _renderCanonicalReportPanel(): React.ReactElement {
    const rows = this._collectCanonicalReport();
    const totalSourcesAffected = new Set(rows.map((r) => r.sourceUrl)).size;
    return (
      <Stack tokens={{ childrenGap: 12 }} style={{ paddingTop: 12 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          Lists every non-canonical SPO document link found in the loaded results.
          A link is non-canonical when it's a sharing-link wrapper, an Office Online
          viewer wrapper, or a direct file URL that isn't in the stable
          {' '}<code style={{ fontFamily: 'Consolas, monospace', fontSize: 11, padding: '1px 4px', background: 'rgba(127,127,127,0.15)', borderRadius: 2 }}>Forms/AllItems.aspx?id=</code> form.
          {' '}<strong>{rows.length.toLocaleString()}</strong> link(s) across{' '}
          <strong>{totalSourcesAffected.toLocaleString()}</strong> source page(s)/file(s).
        </MessageBar>

        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="small" styles={{ root: { opacity: 0.75 } }}>
            Reflects only the currently selected scan.
            {this.state.selectedJob && (
              <> Source: <strong>{this.state.selectedJob.kind === 'documents' ? 'Document' : 'Page'} scan</strong> from {new Date(this.state.selectedJob.startedAt).toLocaleString()}.</>
            )}
          </Text>
          <button
            disabled={rows.length === 0}
            onClick={this._exportCanonicalReportCsv}
            style={{
              padding: '6px 16px',
              background: rows.length > 0 ? '#0078d4' : 'rgba(127,127,127,0.4)',
              color: '#fff',
              border: 'none',
              borderRadius: 2,
              fontSize: 13,
              cursor: rows.length > 0 ? 'pointer' : 'not-allowed',
            }}
          >
            Download CSV
          </button>
        </Stack>

        {rows.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.success}>
            No non-canonical SPO document links found in the loaded results. Everything's
            already in the stable AllItems form, or there's nothing to canonicalize.
          </MessageBar>
        ) : (
          <div
            style={{
              maxHeight: '70vh',
              overflowY: 'scroll',
              border: '1px solid rgba(127,127,127,0.4)',
              borderRadius: 2,
              padding: 4,
              background: 'rgba(127,127,127,0.06)',
            }}
          >
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ borderBottom: '1px solid rgba(127,127,127,0.4)', textAlign: 'left' }}>
                  <th style={{ padding: '6px 8px', width: 60 }}>Kind</th>
                  <th style={{ padding: '6px 8px', minWidth: 180 }}>File / Page</th>
                  <th style={{ padding: '6px 8px', minWidth: 140 }}>Link text</th>
                  <th style={{ padding: '6px 8px', minWidth: 240 }}>Current link</th>
                  <th style={{ padding: '6px 8px', minWidth: 240 }}>Suggested canonical</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const isDoc = r.sourceKind === 'document';
                  const sourceHref = r.sourceUrl.startsWith('http')
                    ? r.sourceUrl
                    : `${TENANT_ORIGIN}${r.sourceUrl}`;
                  return (
                    <tr
                      key={`${r.sourceUrl}::${r.current}::${i}`}
                      style={{ borderBottom: '1px solid rgba(127,127,127,0.2)' }}
                    >
                      <td style={{ padding: '6px 8px', verticalAlign: 'top' }}>
                        <span
                          style={{
                            background: isDoc ? 'rgba(74,144,226,0.18)' : 'rgba(160,116,213,0.18)',
                            color: isDoc ? '#4a90e2' : '#a074d5',
                            border: isDoc ? '1px solid rgba(74,144,226,0.45)' : '1px solid rgba(160,116,213,0.45)',
                            padding: '1px 6px',
                            borderRadius: 8,
                            fontSize: 10,
                            fontWeight: 700,
                            letterSpacing: 0.4,
                            textTransform: 'uppercase',
                          }}
                        >
                          {isDoc ? 'doc' : 'page'}
                        </span>
                      </td>
                      <td style={{ padding: '6px 8px', verticalAlign: 'top' }}>
                        <a
                          href={sourceHref}
                          target="_blank"
                          rel="noopener noreferrer"
                          title={r.sourceUrl}
                          style={{ color: 'inherit', textDecoration: 'underline' }}
                        >
                          {r.source}
                        </a>
                        <div style={{ fontSize: 10, opacity: 0.6, fontFamily: 'Consolas, monospace' }}>{r.site}</div>
                      </td>
                      <td style={{ padding: '6px 8px', verticalAlign: 'top' }}>
                        {r.text || <em style={{ opacity: 0.5 }}>(none)</em>}
                      </td>
                      <td style={{ padding: '6px 8px', verticalAlign: 'top', fontFamily: 'Consolas, monospace', fontSize: 11, wordBreak: 'break-all' }}>
                        <a
                          href={r.current}
                          target="_blank"
                          rel="noopener noreferrer"
                          style={{ color: 'inherit', textDecoration: 'underline' }}
                        >
                          {r.current}
                        </a>
                      </td>
                      <td style={{ padding: '6px 8px', verticalAlign: 'top', fontFamily: 'Consolas, monospace', fontSize: 11, wordBreak: 'break-all' }}>
                        <a
                          href={r.suggested}
                          target="_blank"
                          rel="noopener noreferrer"
                          style={{ color: '#3aaa3a', textDecoration: 'underline' }}
                        >
                          {r.suggested}
                        </a>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Stack>
    );
  }

  private async _openSchedulePanel(): Promise<void> {
    this.setState({ scheduleOpen: true, scheduleLoading: true, scheduleError: '' });
    try {
      const config = await this.props.service.getSchedule();
      this.setState({
        scheduleConfig: config,
        scheduleLoading: false,
        scheduleEnabledDraft: config.enabled,
        scheduleTimeOfDayDraft: config.timeOfDay,
        scheduleTimeZoneDraft: config.timeZone,
      });
    } catch (err) {
      this.setState({
        scheduleLoading: false,
        scheduleError: (err as Error).message,
      });
    }
  }

  private async _saveSchedule(): Promise<void> {
    const { scheduleEnabledDraft, scheduleTimeOfDayDraft, scheduleTimeZoneDraft } = this.state;
    this.setState({ scheduleSaving: true, scheduleError: '' });
    try {
      const next = await this.props.service.setSchedule({
        enabled: scheduleEnabledDraft,
        timeOfDay: scheduleTimeOfDayDraft,
        timeZone: scheduleTimeZoneDraft,
      });
      this.setState({
        scheduleConfig: next,
        scheduleSaving: false,
        scheduleEnabledDraft: next.enabled,
        scheduleTimeOfDayDraft: next.timeOfDay,
        scheduleTimeZoneDraft: next.timeZone,
      });
    } catch (err) {
      this.setState({
        scheduleSaving: false,
        scheduleError: (err as Error).message,
      });
    }
  }

  private _renderSchedulePanel(): React.ReactElement {
    const {
      scheduleConfig,
      scheduleLoading,
      scheduleSaving,
      scheduleError,
      scheduleEnabledDraft,
      scheduleTimeOfDayDraft,
      scheduleTimeZoneDraft,
    } = this.state;

    const dirty =
      !!scheduleConfig &&
      (scheduleEnabledDraft !== scheduleConfig.enabled
        || scheduleTimeOfDayDraft !== scheduleConfig.timeOfDay
        || scheduleTimeZoneDraft !== scheduleConfig.timeZone);

    return (
      <Stack tokens={{ childrenGap: 12 }} style={{ paddingTop: 12 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          The daily scheduled scan runs a paired page + document scan once per day.
          The doc scan reuses unchanged files via SP-reported ETag (no download for
          stable files), and the page scan reuses pages whose <em>Modified</em> matches
          the previous scan. Only changes are extracted; the rest is rolled forward
          from the previous run.
        </MessageBar>

        {scheduleLoading && <Text>Loading schedule…</Text>}

        {scheduleError && (
          <MessageBar messageBarType={MessageBarType.error}>{scheduleError}</MessageBar>
        )}

        {!scheduleLoading && scheduleConfig && (
          <>
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={scheduleEnabledDraft}
                onChange={(e) => this.setState({ scheduleEnabledDraft: e.target.checked })}
              />
              <span style={{ fontSize: 14 }}><strong>Enable daily scheduled scan</strong></span>
            </label>

            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="small">Time of day (24-hour)</Text>
              <input
                type="time"
                value={scheduleTimeOfDayDraft}
                onChange={(e) => this.setState({ scheduleTimeOfDayDraft: e.target.value })}
                disabled={!scheduleEnabledDraft}
                style={{ padding: 6, fontSize: 14, width: 140 }}
              />
            </Stack>

            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="small">Time zone (IANA)</Text>
              <select
                value={scheduleTimeZoneDraft}
                onChange={(e) => this.setState({ scheduleTimeZoneDraft: e.target.value })}
                disabled={!scheduleEnabledDraft}
                style={{ padding: 6, fontSize: 14, width: 240 }}
              >
                <option value="America/Chicago">America/Chicago (Central)</option>
                <option value="America/New_York">America/New_York (Eastern)</option>
                <option value="America/Denver">America/Denver (Mountain)</option>
                <option value="America/Los_Angeles">America/Los_Angeles (Pacific)</option>
                <option value="UTC">UTC</option>
              </select>
            </Stack>

            <Stack tokens={{ childrenGap: 4 }} style={{ marginTop: 8, paddingTop: 8, borderTop: '1px solid rgba(127,127,127,0.3)' }}>
              <Text variant="small"><strong>Last run</strong></Text>
              <Text variant="small">
                {scheduleConfig.lastFiredAt
                  ? `${new Date(scheduleConfig.lastFiredAt).toLocaleString()} (job ${scheduleConfig.lastJobId.slice(0, 8)}…)`
                  : 'Never'}
              </Text>
              <Text variant="small">
                <strong>Next run:</strong> next time the hour ticks past <code>{scheduleTimeOfDayDraft}</code> in <code>{scheduleTimeZoneDraft}</code> (timer polls hourly)
              </Text>
              {scheduleConfig.updatedAt && (
                <Text variant="small" style={{ opacity: 0.7 }}>
                  Last edited: {new Date(scheduleConfig.updatedAt).toLocaleString()} by {scheduleConfig.updatedBy}
                </Text>
              )}
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 12 }}>
              <button
                disabled={!dirty || scheduleSaving}
                onClick={() => { void this._saveSchedule(); }}
                style={{
                  padding: '8px 24px',
                  background: !dirty || scheduleSaving ? 'rgba(127,127,127,0.4)' : '#0078d4',
                  color: '#fff',
                  border: 'none',
                  borderRadius: 2,
                  fontSize: 14,
                  cursor: !dirty || scheduleSaving ? 'not-allowed' : 'pointer',
                }}
              >
                {scheduleSaving ? 'Saving…' : 'Save'}
              </button>
              <button
                onClick={() => this.setState({ scheduleOpen: false })}
                style={{
                  padding: '8px 24px',
                  background: 'transparent',
                  color: 'inherit',
                  border: '1px solid rgba(127,127,127,0.5)',
                  borderRadius: 2,
                  fontSize: 14,
                  cursor: 'pointer',
                }}
              >
                Close
              </button>
            </Stack>
          </>
        )}
      </Stack>
    );
  }

  private _renderPageScanPanel(): React.ReactElement {
    const {
      pageScanSelectedSites,
      pageScanSitePickerFilter,
      pageScanVerifyFiles,
      sitePickerSites,
      sitePickerHubs,
      sitePickerLoading,
      sitePickerError,
      triggering,
    } = this.state;
    return (
      <Stack tokens={{ childrenGap: 12 }} style={{ paddingTop: 12 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          A page scan walks every modern Site Page across the selected sites
          and extracts every URL on each page. Pick sites below, or leave the
          selection empty to scan every site in the tenant.
        </MessageBar>

        {this._renderSitePicker(
          'page',
          sitePickerSites,
          sitePickerHubs,
          sitePickerLoading,
          sitePickerError,
          pageScanSelectedSites,
          pageScanSitePickerFilter,
        )}

        {pageScanSelectedSites.size === 0 && (
          <MessageBar messageBarType={MessageBarType.warning}>
            No sites selected — the scan will cover <strong>every site in the tenant</strong>.
          </MessageBar>
        )}

        <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer' }}>
          <input
            type="checkbox"
            checked={pageScanVerifyFiles}
            onChange={(e) => this.setState({ pageScanVerifyFiles: e.target.checked })}
            style={{ marginTop: 3 }}
          />
          <span style={{ fontSize: 13 }}>
            <strong>Verify SPO file links</strong> — after the scan, HEAD-check every
            <code style={{ margin: '0 4px' }}>AllItems.aspx?id=</code>link against SP REST and
            flag 404s as <em>malformed-spo-link</em>. Catches broken links from bad find/replaces
            that left the URL shape valid but pointing at a non-existent file. Slower — one REST
            call per unique AllItems link.
          </span>
        </label>

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <button
            disabled={triggering}
            onClick={() => { void this._triggerScan(); }}
            style={{
              padding: '8px 24px',
              background: triggering ? 'rgba(127,127,127,0.4)' : '#0078d4',
              color: '#fff',
              border: 'none',
              borderRadius: 2,
              fontSize: 14,
              cursor: triggering ? 'not-allowed' : 'pointer',
            }}
          >
            {triggering ? 'Starting...' : 'Start scan'}
          </button>
          <button
            onClick={() => this.setState({ pageScanOpen: false })}
            style={{
              padding: '8px 24px',
              background: 'transparent',
              color: 'inherit',
              border: '1px solid rgba(127,127,127,0.5)',
              borderRadius: 2,
              fontSize: 14,
              cursor: 'pointer',
            }}
          >
            Cancel
          </button>
        </Stack>
      </Stack>
    );
  }

  private _renderDocScanPanel(): React.ReactElement {
    const {
      docScanMaxMB,
      docScanFileRefs,
      docScanModifiedAfter,
      docScanPreviewOnly,
      docScanVerifyFiles,
      docScanSelectedSites,
      docScanSitePickerFilter,
      sitePickerSites,
      sitePickerHubs,
      sitePickerLoading,
      sitePickerError,
      triggering,
    } = this.state;
    const targetingFiles = docScanFileRefs.trim().length > 0;
    return (
      <Stack tokens={{ childrenGap: 12 }} style={{ paddingTop: 12 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          A document scan walks every modern document library in scope, downloads
          each <strong>.docx / .xlsx / .pptx / .pdf</strong> file under the size limit,
          and extracts every embedded URL. Other file types are skipped.
        </MessageBar>

        <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="end">
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>
              Max file size (MB)
            </Text>
            <input
              type="number"
              min={1}
              max={500}
              step={10}
              value={docScanMaxMB}
              onChange={(e) => {
                const v = parseInt(e.target.value, 10);
                this.setState({ docScanMaxMB: Number.isNaN(v) ? 100 : v });
              }}
              style={{ ...this._inputStyle(false), width: 100, fontSize: 14 }}
            />
          </Stack>
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>
              Modified after <span style={{ fontStyle: 'italic', opacity: 0.6 }}>(optional)</span>
            </Text>
            <input
              type="date"
              value={docScanModifiedAfter}
              onChange={(e) => this.setState({ docScanModifiedAfter: e.target.value })}
              style={{ ...this._inputStyle(false), width: 160, fontSize: 14 }}
            />
          </Stack>
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, paddingBottom: 6, cursor: 'pointer' }}>
            <input
              type="checkbox"
              checked={docScanPreviewOnly}
              onChange={(e) => this.setState({ docScanPreviewOnly: e.target.checked })}
            />
            <Text>Preview only (count files, don't scan)</Text>
          </label>
        </Stack>

        {!targetingFiles && this._renderSitePicker(
          'doc',
          sitePickerSites,
          sitePickerHubs,
          sitePickerLoading,
          sitePickerError,
          docScanSelectedSites,
          docScanSitePickerFilter,
        )}

        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>
            Targeted files <span style={{ fontStyle: 'italic', opacity: 0.6 }}>(optional, one per line)</span>
          </Text>
          <textarea
            value={docScanFileRefs}
            onChange={(e) => this.setState({ docScanFileRefs: e.target.value })}
            placeholder={'/sites/hub/Shared Documents/Annual Report 2025.pdf\n/sites/online-training-manual/Shared Documents/Onboarding.docx'}
            style={this._textareaStyle(80)}
          />
          <Text variant="small" styles={{ root: { opacity: 0.75 } }}>
            Server-relative file paths to scan directly (skips library enumeration).
            Useful for one-off rescans of specific files. Overrides the site selection above.
          </Text>
        </Stack>

        {!targetingFiles && docScanSelectedSites.size === 0 && (
          <MessageBar messageBarType={MessageBarType.warning}>
            No sites selected — the scan will cover <strong>every site in the tenant</strong>.
            {!docScanPreviewOnly && ' Consider running a Preview first to see how many files are involved.'}
          </MessageBar>
        )}

        <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer' }}>
          <input
            type="checkbox"
            checked={docScanVerifyFiles}
            onChange={(e) => this.setState({ docScanVerifyFiles: e.target.checked })}
            style={{ marginTop: 3 }}
          />
          <span style={{ fontSize: 13 }}>
            <strong>Verify SPO file links</strong> — after the scan, HEAD-check every
            <code style={{ margin: '0 4px' }}>AllItems.aspx?id=</code>link against SP REST and
            flag 404s as <em>malformed-spo-link</em>. Slower — one REST call per unique AllItems link.
          </span>
        </label>

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <button
            disabled={triggering}
            onClick={() => { void this._triggerDocScan(); }}
            style={{
              padding: '8px 24px',
              background: triggering ? 'rgba(127,127,127,0.4)' : '#0078d4',
              color: '#fff',
              border: 'none',
              borderRadius: 2,
              fontSize: 14,
              cursor: triggering ? 'not-allowed' : 'pointer',
            }}
          >
            {triggering ? 'Starting...' : (docScanPreviewOnly ? 'Run preview' : 'Start scan')}
          </button>
          <button
            onClick={() => this.setState({ docScanOpen: false })}
            style={{
              padding: '8px 24px',
              background: 'transparent',
              color: 'inherit',
              border: '1px solid rgba(127,127,127,0.5)',
              borderRadius: 2,
              fontSize: 14,
              cursor: 'pointer',
            }}
          >
            Cancel
          </button>
        </Stack>
      </Stack>
    );
  }

  /**
   * Site picker checklist — shared by both the doc scan and page scan
   * dialogs. Sites are grouped by hub (with a "Standalone" group for
   * sites that aren't associated with any hub). Each hub group has a
   * header with a "select all in hub" toggle.
   *
   * `target` is `'doc'` or `'page'` so the toggle methods know which
   * piece of state to mutate. Same picker, two consumers.
   */
  private _renderSitePicker(
    target: 'doc' | 'page',
    sites: ISiteSummary[],
    hubs: IHubSummary[],
    loading: boolean,
    error: string,
    selected: Set<string>,
    filter: string,
  ): React.ReactElement {
    if (loading) {
      return <Spinner size={SpinnerSize.medium} label="Loading sites..." />;
    }
    if (error) {
      return (
        <MessageBar messageBarType={MessageBarType.error}>
          Site list failed to load: {error}
        </MessageBar>
      );
    }
    if (sites.length === 0) {
      return (
        <MessageBar messageBarType={MessageBarType.info}>
          No sites loaded yet. Open the doc scan panel to fetch the tenant site list.
        </MessageBar>
      );
    }

    const filterLower = filter.trim().toLowerCase();
    const matches = (s: ISiteSummary): boolean => {
      if (!filterLower) return true;
      return (
        s.title.toLowerCase().indexOf(filterLower) !== -1 ||
        s.serverRelativeUrl.toLowerCase().indexOf(filterLower) !== -1
      );
    };

    // Group filtered sites by hub. Sites without a hub go into "Standalone".
    const STANDALONE = '__standalone__';
    const groups = new Map<string, ISiteSummary[]>();
    for (const s of sites) {
      if (!matches(s)) continue;
      const key = s.hubSiteId ?? STANDALONE;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key)!.push(s);
    }
    groups.forEach((arr) => arr.sort((a, b) => a.title.localeCompare(b.title)));

    // Build group display order: hubs first (alphabetical), then standalone.
    const orderedGroupKeys: string[] = [];
    for (const h of hubs) {
      if (groups.has(h.id)) orderedGroupKeys.push(h.id);
    }
    if (groups.has(STANDALONE)) orderedGroupKeys.push(STANDALONE);

    const filteredCount = Array.from(groups.values()).reduce((s, g) => s + g.length, 0);

    return (
      <Stack tokens={{ childrenGap: 6 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>
            Sites to scan ({selected.size} selected · {filteredCount} shown · {sites.length} total)
          </Text>
          <Stack horizontal tokens={{ childrenGap: 6 }}>
            <button
              onClick={() => this._toggleSelectAllSites(target, true)}
              style={this._smallButtonStyle()}
              title="Select every visible site"
            >
              All visible
            </button>
            <button
              onClick={() => this._toggleSelectAllSites(target, false)}
              style={this._smallButtonStyle()}
              title="Clear selection"
            >
              None
            </button>
          </Stack>
        </Stack>
        <input
          type="text"
          value={filter}
          onChange={(e) => {
            const v = e.target.value;
            if (target === 'doc') this.setState({ docScanSitePickerFilter: v });
            else this.setState({ pageScanSitePickerFilter: v });
          }}
          placeholder="Filter sites by title or path..."
          style={{ ...this._inputStyle(false), fontSize: 13 }}
        />
        <div
          // Mobile + nested-scroll friendliness:
          //   - `maxHeight: 50vh` (not fixed px) so the picker scales
          //     with the viewport — on a phone, 420px would be most of
          //     the screen and the user could only scroll the outer
          //     Fluent Panel, not this inner box.
          //   - `overscrollBehavior: contain` stops the inner scroll
          //     from "chaining" up to the outer Panel when it hits a
          //     boundary, which on mobile is the typical reason a
          //     nested scroll container feels broken.
          //   - `WebkitOverflowScrolling: touch` enables iOS momentum
          //     scrolling inside the box.
          //   - `touchAction: pan-y` declares vertical pan as the
          //     intended gesture so the browser's touch-handler picks
          //     it up instead of treating it as a click.
          //   - `overflowY: scroll` (not `auto`) so the scrollbar is
          //     always visible on desktop dark themes too.
          //   - `scrollbarGutter: stable` reserves the gutter so the
          //     layout doesn't jump when content first overflows.
          style={{
            maxHeight: '50vh',
            minHeight: 240,
            overflowY: 'scroll',
            scrollbarGutter: 'stable',
            overscrollBehavior: 'contain',
            WebkitOverflowScrolling: 'touch',
            touchAction: 'pan-y',
            border: '1px solid rgba(127,127,127,0.4)',
            borderRadius: 2,
            padding: 8,
            background: 'rgba(127,127,127,0.06)',
          }}
        >
          {orderedGroupKeys.length === 0 && (
            <Text variant="small" styles={{ root: { opacity: 0.7, fontStyle: 'italic' } }}>
              No sites match the filter.
            </Text>
          )}
          {orderedGroupKeys.map((key) => {
            const groupSites = groups.get(key)!;
            const hub = hubs.find((h) => h.id === key);
            const groupLabel = hub ? `${hub.title} (hub)` : 'Standalone sites';
            const allSelected = groupSites.every((s) => selected.has(s.serverRelativeUrl));
            return (
              <div key={key} style={{ marginBottom: 12 }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginBottom: 4 }}>
                  <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                    {groupLabel}
                  </Text>
                  <button
                    onClick={() => this._toggleSelectGroup(target, groupSites, !allSelected)}
                    style={this._smallButtonStyle()}
                    title={allSelected ? 'Deselect all sites in this group' : 'Select all sites in this group'}
                  >
                    {allSelected ? 'Clear group' : 'Select all'}
                  </button>
                </Stack>
                {groupSites.map((s) => {
                  const isSelected = selected.has(s.serverRelativeUrl);
                  return (
                    <label
                      key={s.serverRelativeUrl}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: 8,
                        padding: '3px 4px',
                        cursor: 'pointer',
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={() => this._toggleSelectSite(target, s.serverRelativeUrl)}
                      />
                      <span style={{ flex: 1, minWidth: 0, display: 'flex', flexDirection: 'column' }}>
                        <span style={{ fontSize: 13, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {s.title || s.serverRelativeUrl}
                        </span>
                        <span style={{ fontSize: 11, opacity: 0.65, fontFamily: 'Consolas, monospace' }}>
                          {s.serverRelativeUrl}
                        </span>
                      </span>
                    </label>
                  );
                })}
              </div>
            );
          })}
        </div>
      </Stack>
    );
  }

  private _smallButtonStyle(): React.CSSProperties {
    return {
      padding: '3px 10px',
      background: 'transparent',
      color: 'inherit',
      border: '1px solid rgba(127,127,127,0.45)',
      borderRadius: 2,
      fontSize: 11,
      cursor: 'pointer',
    };
  }

  private _toggleSelectSite = (target: 'doc' | 'page', path: string): void => {
    const current = target === 'doc'
      ? this.state.docScanSelectedSites
      : this.state.pageScanSelectedSites;
    const next = new Set(current);
    if (next.has(path)) next.delete(path);
    else next.add(path);
    if (target === 'doc') this.setState({ docScanSelectedSites: next });
    else this.setState({ pageScanSelectedSites: next });
  };

  private _toggleSelectGroup = (
    target: 'doc' | 'page',
    sites: ISiteSummary[],
    select: boolean,
  ): void => {
    const current = target === 'doc'
      ? this.state.docScanSelectedSites
      : this.state.pageScanSelectedSites;
    const next = new Set(current);
    for (const s of sites) {
      if (select) next.add(s.serverRelativeUrl);
      else next.delete(s.serverRelativeUrl);
    }
    if (target === 'doc') this.setState({ docScanSelectedSites: next });
    else this.setState({ pageScanSelectedSites: next });
  };

  private _toggleSelectAllSites = (target: 'doc' | 'page', select: boolean): void => {
    if (!select) {
      if (target === 'doc') this.setState({ docScanSelectedSites: new Set() });
      else this.setState({ pageScanSelectedSites: new Set() });
      return;
    }
    // Select all currently visible (i.e., matching the filter) sites.
    const filter = target === 'doc'
      ? this.state.docScanSitePickerFilter
      : this.state.pageScanSitePickerFilter;
    const filterLower = filter.trim().toLowerCase();
    const next = new Set<string>();
    for (const s of this.state.sitePickerSites) {
      const visible = !filterLower ||
        s.title.toLowerCase().indexOf(filterLower) !== -1 ||
        s.serverRelativeUrl.toLowerCase().indexOf(filterLower) !== -1;
      if (visible) next.add(s.serverRelativeUrl);
    }
    if (target === 'doc') this.setState({ docScanSelectedSites: next });
    else this.setState({ pageScanSelectedSites: next });
  };

  /**
   * Render the preview-only doc scan results: per-site/per-library
   * file counts with checkboxes to select which sites to promote to a
   * real scan.
   */
  private _renderPreviewView(preview: IPreviewResults): React.ReactElement {
    const { previewSelectedSites, promoting, isAdmin } = this.state;
    const totalBytesMb = (preview.totals.bytes / 1024 / 1024).toFixed(1);

    // Compute selected counts for the bar
    let selectedFiles = 0;
    let selectedBytes = 0;
    for (const s of preview.siteSummaries) {
      if (previewSelectedSites.has(s.site)) {
        selectedFiles += s.totalFiles;
        selectedBytes += s.totalBytes;
      }
    }
    const selectedMb = (selectedBytes / 1024 / 1024).toFixed(1);

    const renderBucket = (
      label: string,
      note: string,
      bucket: IPreviewBucketStats,
    ): React.ReactElement => {
      const topExt = Object.keys(bucket.byExtension)
        .map((ext) => ({ ext, stats: bucket.byExtension[ext] }))
        .sort((a, b) => b.stats.count - a.stats.count)
        .slice(0, 8);
      const mb = (bucket.totalBytes / 1024 / 1024).toFixed(1);
      return (
        <div
          style={{
            flex: 1,
            padding: 10,
            border: '1px solid rgba(127,127,127,0.3)',
            borderRadius: 2,
            background: 'rgba(127,127,127,0.04)',
          }}
        >
          <div style={{ fontWeight: 600, marginBottom: 4 }}>{label}</div>
          <div style={{ fontSize: 12, opacity: 0.8, marginBottom: 6 }}>{note}</div>
          <div style={{ fontSize: 12 }}>
            <strong>{bucket.files.toLocaleString()}</strong> files · {mb} MB
          </div>
          {topExt.length > 0 && (
            <div style={{ marginTop: 6, fontSize: 11, opacity: 0.9 }}>
              {topExt.map(({ ext, stats }) => (
                <span
                  key={ext}
                  style={{
                    display: 'inline-block',
                    padding: '1px 6px',
                    margin: '2px 4px 2px 0',
                    border: '1px solid rgba(127,127,127,0.3)',
                    borderRadius: 10,
                    fontFamily: 'Consolas, monospace',
                  }}
                >
                  .{ext || '(none)'} · {stats.count.toLocaleString()}
                </span>
              ))}
            </div>
          )}
        </div>
      );
    };

    return (
      <Stack tokens={{ childrenGap: 10 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          <strong>Preview only.</strong> This scan enumerated files but didn't open any of them.
          Tick the sites you want to actually scan, then click <em>Promote selected to scan</em>.
          Total: <strong>{preview.totals.files.toLocaleString()}</strong> files
          ({totalBytesMb} MB) across <strong>{preview.totals.sites}</strong> sites.
        </MessageBar>

        {preview.included && preview.excluded && (
          <Stack horizontal tokens={{ childrenGap: 10 }}>
            {renderBucket(
              'Included (will be scanned)',
              'Supported for link extraction — .docx / .xlsx / .pptx / .pdf and OOXML variants.',
              preview.included,
            )}
            {renderBucket(
              'Excluded (today)',
              'Not scanned for links. Available for duplicate detection once opt-in is wired.',
              preview.excluded,
            )}
          </Stack>
        )}

        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="small">
            <strong>{previewSelectedSites.size}</strong> of {preview.totals.sites} sites selected
            {' · '}
            <strong>{selectedFiles.toLocaleString()}</strong> files
            {' · '}
            {selectedMb} MB
          </Text>
          <Stack horizontal tokens={{ childrenGap: 6 }}>
            <button
              onClick={() => this.setState({
                previewSelectedSites: new Set(preview.siteSummaries.map((s) => s.site)),
              })}
              style={this._smallButtonStyle()}
            >
              Select all
            </button>
            <button
              onClick={() => this.setState({ previewSelectedSites: new Set() })}
              style={this._smallButtonStyle()}
            >
              None
            </button>
            <button
              disabled={!isAdmin || promoting || previewSelectedSites.size === 0}
              onClick={() => { void this._promotePreview(); }}
              style={{
                padding: '6px 16px',
                background: (!isAdmin || promoting || previewSelectedSites.size === 0)
                  ? 'rgba(127,127,127,0.4)'
                  : '#107c10',
                color: '#fff',
                border: 'none',
                borderRadius: 2,
                fontSize: 13,
                cursor: (!isAdmin || promoting || previewSelectedSites.size === 0)
                  ? 'not-allowed'
                  : 'pointer',
              }}
              title={isAdmin
                ? 'Start a real scan against the selected sites'
                : 'Admin only'}
            >
              {promoting ? 'Starting...' : 'Promote selected to scan'}
            </button>
          </Stack>
        </Stack>

        <div
          style={{
            maxHeight: 480,
            overflowY: 'auto',
            border: '1px solid rgba(127,127,127,0.4)',
            borderRadius: 2,
            background: 'rgba(127,127,127,0.06)',
          }}
        >
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr style={{ borderBottom: '1px solid rgba(127,127,127,0.4)', textAlign: 'left' }}>
                <th style={{ padding: '6px 8px', width: 32 }}></th>
                <th style={{ padding: '6px 8px' }}>Site</th>
                <th style={{ padding: '6px 8px', width: 100, textAlign: 'right' }}>Files</th>
                <th style={{ padding: '6px 8px', width: 110, textAlign: 'right' }}>Size</th>
                <th style={{ padding: '6px 8px' }}>Libraries</th>
              </tr>
            </thead>
            <tbody>
              {preview.siteSummaries.map((s) => {
                const selected = previewSelectedSites.has(s.site);
                const sizeMb = (s.totalBytes / 1024 / 1024).toFixed(1);
                const libNames = Object.keys(s.libraries).sort().map((k) => `${k} (${s.libraries[k].files})`).join(', ');
                return (
                  <tr
                    key={s.site}
                    style={{
                      borderBottom: '1px solid rgba(127,127,127,0.2)',
                      cursor: 'pointer',
                      background: selected ? 'rgba(58,150,221,0.08)' : undefined,
                    }}
                    onClick={() => this._togglePreviewSite(s.site)}
                  >
                    <td style={{ padding: '6px 8px' }}>
                      <input
                        type="checkbox"
                        checked={selected}
                        onChange={() => this._togglePreviewSite(s.site)}
                        onClick={(e) => e.stopPropagation()}
                      />
                    </td>
                    <td style={{ padding: '6px 8px', fontFamily: 'Consolas, monospace' }}>{s.site}</td>
                    <td style={{ padding: '6px 8px', textAlign: 'right' }}>{s.totalFiles.toLocaleString()}</td>
                    <td style={{ padding: '6px 8px', textAlign: 'right' }}>{sizeMb} MB</td>
                    <td style={{ padding: '6px 8px', opacity: 0.8 }}>{libNames}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Stack>
    );
  }

  private _togglePreviewSite = (site: string): void => {
    const next = new Set(this.state.previewSelectedSites);
    if (next.has(site)) next.delete(site);
    else next.add(site);
    this.setState({ previewSelectedSites: next });
  };

  private _promotePreview = async (): Promise<void> => {
    const { selectedJobId, previewSelectedSites } = this.state;
    if (!selectedJobId || previewSelectedSites.size === 0) return;
    this.setState({ promoting: true });
    try {
      const sites = Array.from(previewSelectedSites);
      const res = await this.props.service.promoteDocScan(selectedJobId, sites);
      // Refresh job list and select the newly created real scan job.
      const jobs = await this.props.service.listJobs();
      this.setState({
        promoting: false,
        jobs,
        selectedJobId: res.jobId,
        selectedJob: jobs.find((j) => j.jobId === res.jobId),
        results: undefined,
        previewResults: undefined,
        previewSelectedSites: new Set(),
      });
      this._startPolling(res.jobId);
    } catch (err) {
      this.setState({ promoting: false });
      this.props.onError(`Promote preview: ${(err as Error).message}`);
    }
  };

  private _loadSitePicker = async (): Promise<void> => {
    if (this.state.sitePickerSites.length > 0) return; // already loaded
    this.setState({ sitePickerLoading: true, sitePickerError: '' });
    try {
      const res = await this.props.service.listSites();
      this.setState({
        sitePickerSites: res.sites,
        sitePickerHubs: res.hubs,
        sitePickerLoading: false,
      });
    } catch (err) {
      this.setState({ sitePickerLoading: false, sitePickerError: (err as Error).message });
    }
  };

  private _renderFindReplacePanel(): React.ReactElement {
    const { frFind, frReplace, frMode, frModeLabel, results, frBusy, frResponse, frApplied, frError, isAdmin } = this.state;
    if (!results) return <Text>No results loaded.</Text>;
    const isCanonical = frMode === 'canonical';

    const find = frFind.trim();

    // Compute client-side preview matches purely for the count display.
    // The authoritative preview comes from the server's dry-run response.
    const clientMatches: ILinkRow[] = [];
    const jobIdForRows = this.state.selectedJobId ?? '';
    const fallbackKind: 'pages' | 'documents' = this.state.selectedJob?.kind ?? 'pages';
    if (find) {
      const lower = find.toLowerCase();
      for (const site of results.sites) {
        for (const page of site.pages) {
          const pageKind: 'pages' | 'documents' = page.sourceKind ?? fallbackKind;
          for (let i = 0; i < page.links.length; i++) {
            const link = page.links[i];
            if (link.url.toLowerCase().indexOf(lower) !== -1) {
              clientMatches.push(
                this._makeRow(site.site, page, link, i, jobIdForRows, pageKind),
              );
            }
          }
        }
      }
    }
    const affectedPageCount = new Set(clientMatches.map((m) => m.page.pageId)).size;

    const canPreview = !!find && !frBusy;
    const canApply = !!frResponse && frResponse.summary.preview > 0 && !frBusy && isAdmin;

    return (
      <Stack tokens={{ childrenGap: 12 }} style={{ paddingTop: 12 }}>
        {!isAdmin && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Find &amp; Replace is admin-only. You can preview matches in the table on the main tab,
            but the Apply action requires SP Redirect Manager Admin group membership.
          </MessageBar>
        )}

        {isCanonical && (
          <MessageBar messageBarType={MessageBarType.info}>
            <strong>{frModeLabel || 'Canonical mode'}.</strong> Matching by <em>canonical key</em>
            instead of substring — every URL form (sharing wrapper, direct path, AllItems URL)
            that resolves to the same target will be replaced with the canonical URL below.
          </MessageBar>
        )}

        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>
            {isCanonical ? 'Canonical key (matched exactly)' : 'Find URLs containing'}
          </Text>
          <input
            value={frFind}
            onChange={(e) => this.setState({ frFind: e.target.value, frResponse: undefined, frApplied: false })}
            placeholder={isCanonical ? '(set automatically by Align All)' : 'e.g. legacy.example.com or /:u:/g/'}
            disabled={isCanonical}
            style={this._inputStyle(isCanonical)}
          />
        </Stack>
        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="smallPlus" styles={{ root: { opacity: 0.75 } }}>Replace with</Text>
          <input
            value={frReplace}
            onChange={(e) => this.setState({ frReplace: e.target.value, frResponse: undefined, frApplied: false })}
            placeholder="e.g. contoso.sharepoint.com"
            style={this._inputStyle(false)}
          />
        </Stack>

        {find && !frResponse && (
          <Text variant="small" styles={{ root: { opacity: 0.75 } }}>
            Client-side filter: <strong>{clientMatches.length}</strong> link(s) across{' '}
            <strong>{affectedPageCount}</strong> page(s) match. Click <em>Preview</em> below to verify
            against the latest scan and check write permissions.
          </Text>
        )}

        {frError && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ frError: '' })}>
            {frError}
          </MessageBar>
        )}

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <button
            disabled={!canPreview}
            onClick={() => { void this._runReplace(true); }}
            style={{
              padding: '8px 24px',
              background: canPreview ? '#0078d4' : 'rgba(127,127,127,0.4)',
              color: '#fff',
              border: 'none',
              borderRadius: 2,
              fontSize: 14,
              cursor: canPreview ? 'pointer' : 'not-allowed',
            }}
          >
            {frBusy && !frApplied ? 'Previewing...' : 'Preview'}
          </button>
          <TooltipHost
            content={
              !isAdmin
                ? 'Apply requires SharePoint Redirect Manager Admin group membership.'
                : !frResponse
                ? 'Run Preview first.'
                : frResponse.summary.preview === 0
                ? 'No pages would change.'
                : 'Apply the change to all preview pages.'
            }
          >
            <span>
              <button
                disabled={!canApply}
                onClick={() => { void this._runReplace(false); }}
                style={{
                  padding: '8px 24px',
                  background: canApply ? '#a4262c' : 'rgba(127,127,127,0.4)',
                  color: '#fff',
                  border: 'none',
                  borderRadius: 2,
                  fontSize: 14,
                  cursor: canApply ? 'pointer' : 'not-allowed',
                }}
              >
                {frBusy && !frResponse ? '...' : 'Apply'}
              </button>
            </span>
          </TooltipHost>
        </Stack>

        {frBusy && <Spinner size={SpinnerSize.small} label="Working..." />}

        {frResponse && this._renderReplaceResponse(frResponse)}

        {frApplied && (
          <MessageBar messageBarType={MessageBarType.success}>
            Applied. {frResponse?.summary.applied ?? 0} page(s) updated.
          </MessageBar>
        )}
      </Stack>
    );
  }

  private _renderReplaceResponse(response: ILinkInventoryReplaceResponse): React.ReactElement {
    const { summary, results, dryRun, droppedSites } = response;
    return (
      <Stack tokens={{ childrenGap: 8 }}>
        <Text>
          <strong>{dryRun ? 'Preview' : 'Result'}:</strong>{' '}
          {summary.totalPages} page(s) considered ·{' '}
          {dryRun ? `${summary.preview} would change` : `${summary.applied} applied`}
          {summary.stale > 0 && ` · ${summary.stale} stale (rescan needed)`}
          {summary.conflict > 0 && ` · ${summary.conflict} conflict`}
          {summary.noMatch > 0 && ` · ${summary.noMatch} no-match`}
          {summary.error > 0 && ` · ${summary.error} error`}
        </Text>

        {droppedSites && droppedSites.length > 0 && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Skipped {droppedSites.length} site(s) — you don't have edit access:{' '}
            {droppedSites.join(', ')}
          </MessageBar>
        )}

        <div style={{
          maxHeight: 320,
          overflow: 'auto',
          border: '1px solid rgba(127,127,127,0.4)',
          borderRadius: 2,
          padding: 8,
          background: 'rgba(127,127,127,0.06)',
        }}>
          {results.map((r) => this._renderReplaceRow(r))}
          {results.length === 0 && (
            <Text variant="small" styles={{ root: { fontStyle: 'italic' } }}>
              No pages matched the find pattern (after permission filtering).
            </Text>
          )}
        </div>
      </Stack>
    );
  }

  private _renderReplaceRow(r: IReplacePageResult): React.ReactElement {
    const statusColor: Record<string, string> = {
      applied: '#107c10',
      preview: '#005a9e',
      stale: '#8a6900',
      conflict: '#a4262c',
      'no-match': '#605e5c',
      'not-found': '#a4262c',
      error: '#a4262c',
    };
    return (
      <div
        key={`${r.site}:${r.pageId}`}
        style={{ marginBottom: 10, paddingBottom: 8, borderBottom: '1px solid rgba(127,127,127,0.25)' }}
      >
        <div style={{ fontSize: 12 }}>
          <span
            style={{
              display: 'inline-block',
              padding: '1px 6px',
              borderRadius: 10,
              background: 'rgba(127,127,127,0.18)',
              color: statusColor[r.status] ?? '#999',
              fontWeight: 600,
              fontSize: 10,
              marginRight: 6,
              textTransform: 'uppercase',
            }}
          >
            {r.status}
          </span>
          <a
            // Prefer the page's direct URL (set by the writer from
            // FileLeafRef). Older clients/servers without that field
            // fall back to the modern library landing page — never the
            // classic `/SitePages/Forms/AllItems.aspx` route, which
            // 404s on modern SP.
            href={`${TENANT_ORIGIN}${r.pageUrl ?? `${r.site}/SitePages`}`}
            target="_blank"
            rel="noopener noreferrer"
            style={{ color: 'inherit', textDecoration: 'underline', fontWeight: 600 }}
            title={r.pageUrl ? `Open ${r.pageTitle ?? 'page'}` : `Open Site Pages on ${r.site}`}
          >
            {r.pageTitle ?? `#${r.pageId}`}
          </a>
          <span style={{ opacity: 0.7 }}> · {r.site}</span>
        </div>
        {r.error && (
          <div style={{ fontSize: 11, color: '#a4262c', marginTop: 2 }}>{r.error}</div>
        )}
        {r.details.slice(0, 5).map((d, i) => (
          <div
            key={i}
            style={{
              fontFamily: 'Consolas, monospace',
              fontSize: 11,
              marginTop: 4,
              opacity: d.matched ? 1 : 0.5,
            }}
          >
            <div>− {d.oldUrl}</div>
            <div style={{ color: '#107c10' }}>+ {d.newUrl}</div>
          </div>
        ))}
        {r.details.length > 5 && (
          <Text variant="small" styles={{ root: { color: '#605e5c', fontStyle: 'italic' } }}>
            ...and {r.details.length - 5} more on this page.
          </Text>
        )}
      </div>
    );
  }
}
