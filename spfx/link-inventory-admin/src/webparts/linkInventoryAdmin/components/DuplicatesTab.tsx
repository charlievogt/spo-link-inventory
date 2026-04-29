import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import {
  LinkInventoryService,
  type IBacklinksResultEntry,
  type IDuplicatesAllowlist,
  type IDuplicatesBootstrapJob,
  type IDuplicatesReport,
} from '../services/LinkInventoryService';

interface IDuplicatesTabProps {
  service: LinkInventoryService;
  /** Tenant origin, e.g. "https://contoso.sharepoint.com". Used to convert
   *  server-relative file paths into clickable absolute URLs. */
  tenantOrigin: string;
  onError: (message: string) => void;
}

interface IDuplicatesTabState {
  loading: boolean;
  report?: IDuplicatesReport;
  allowlist?: IDuplicatesAllowlist;
  isAdmin: boolean;
  exporting: boolean;
  addKind: 'hash' | 'path' | 'name';
  addValue: string;
  addNote: string;
  mutating: boolean;
  bootstrapSitePath: string;
  bootstrapLibraryTitle: string;
  bootstrapMaxFiles: string;
  bootstrapMaxVersions: string;
  bootstrapIncludeOther: boolean;
  bootstrapRunning: boolean;
  /** Currently-tracked bootstrap job (polled while running). */
  bootstrapActiveJob?: IDuplicatesBootstrapJob;
  /** fileRef → backlinks (visible entries + hidden count). Loaded after the report. */
  backlinksByRef: Record<string, IBacklinksResultEntry>;
  backlinksLoading: boolean;
  expandedBacklinks: Set<string>;
}

function toAbs(tenantOrigin: string, p: string): string {
  if (/^https?:\/\//i.test(p)) return p;
  return tenantOrigin + (p.startsWith('/') ? p : '/' + p);
}

function fmtBytes(b: number): string {
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / 1024 / 1024).toFixed(1)} MB`;
}

/**
 * Shared props for every cross-page anchor in this tab. SharePoint's
 * modern page router intercepts same-origin clicks and navigates
 * in-place, which silently clobbers `target="_blank"`. The documented
 * escape hatch is `data-interception="off"` (see MS Learn
 * "Hyperlinking considerations in SharePoint Framework"), which tells
 * the router to let the browser handle the click normally.
 */
const EXTERNAL_LINK_ATTRS = {
  target: '_blank',
  rel: 'noopener noreferrer',
  'data-interception': 'off',
} as const;

/** Theme-aware anchor styles — primary link color on dark/light SP themes. */
const LINK_STYLE: React.CSSProperties = {
  color: 'var(--communicationPrimary, #3a96dd)',
  textDecoration: 'none',
};
const LINK_STYLE_MONO: React.CSSProperties = {
  ...LINK_STYLE,
  fontFamily: 'Consolas, monospace',
};

/**
 * Derive the library root URL from a fileRef + sitePath. The library
 * title (`f.library`) can differ from its URL segment ("Documents" title
 * for "Shared Documents" folder), so we take the first path segment
 * after the site root instead of trusting the title.
 */
function libraryUrlFromFileRef(tenantOrigin: string, sitePath: string, fileRef: string): string | undefined {
  if (!fileRef || !sitePath) return undefined;
  const site = sitePath.toLowerCase();
  const ref = fileRef.toLowerCase();
  if (!ref.startsWith(site + '/')) return undefined;
  const remainder = fileRef.slice(sitePath.length + 1);
  const firstSegment = remainder.split('/', 1)[0];
  if (!firstSegment) return undefined;
  return toAbs(tenantOrigin, sitePath + '/' + firstSegment);
}

export class DuplicatesTab extends React.Component<IDuplicatesTabProps, IDuplicatesTabState> {
  private _mounted = false;

  constructor(props: IDuplicatesTabProps) {
    super(props);
    this.state = {
      loading: true,
      isAdmin: false,
      exporting: false,
      addKind: 'hash',
      addValue: '',
      addNote: '',
      mutating: false,
      bootstrapSitePath: '',
      bootstrapLibraryTitle: 'Documents',
      bootstrapMaxFiles: '500',
      bootstrapMaxVersions: '20',
      bootstrapIncludeOther: false,
      bootstrapRunning: false,
      backlinksByRef: {},
      backlinksLoading: false,
      expandedBacklinks: new Set<string>(),
    };
  }

  public componentDidMount(): void {
    this._mounted = true;
    void this._load();
  }

  private _bootstrapPollHandle: number | undefined;

  public componentWillUnmount(): void {
    this._mounted = false;
    if (this._bootstrapPollHandle !== undefined) {
      window.clearInterval(this._bootstrapPollHandle);
      this._bootstrapPollHandle = undefined;
    }
  }

  /** Safe setState that no-ops after unmount — guards against async callbacks racing navigation. */
  private _safeSetState = <K extends keyof IDuplicatesTabState>(
    patch: Pick<IDuplicatesTabState, K>,
  ): void => {
    if (!this._mounted) return;
    this.setState(patch);
  };

  private _load = async (): Promise<void> => {
    this._safeSetState({ loading: true });
    try {
      const [report, al] = await Promise.all([
        this.props.service.getDuplicatesReport(),
        this.props.service.getDuplicatesAllowlist(),
      ]);
      this._safeSetState({
        loading: false,
        report,
        allowlist: al.allowlist,
        isAdmin: report.isAdmin || al.isAdmin,
      });
      // Fire and forget — backlinks render incrementally below each row;
      // we don't block the report on them.
      void this._loadBacklinks(report);
    } catch (err) {
      this._safeSetState({ loading: false });
      this.props.onError(`Load duplicates: ${(err as Error).message}`);
    }
  };

  private _loadBacklinks = async (report: IDuplicatesReport): Promise<void> => {
    const refs = new Set<string>();
    for (const g of report.exactGroups) {
      for (const f of g.files) refs.add(f.fileRef);
    }
    for (const p of report.stalePairs) {
      refs.add(p.staleFileRef);
      refs.add(p.authoritativeFileRef);
    }
    for (const p of report.sameNamePairs) {
      refs.add(p.aFileRef);
      refs.add(p.bFileRef);
    }
    if (refs.size === 0) return;
    this._safeSetState({ backlinksLoading: true });
    try {
      const res = await this.props.service.getBacklinksBatch(Array.from(refs));
      this._safeSetState({ backlinksByRef: res.results, backlinksLoading: false });
    } catch (err) {
      this._safeSetState({ backlinksLoading: false });
      // Non-fatal — duplicates report still renders without backlinks.
      // eslint-disable-next-line no-console
      console.warn(`Backlinks batch load failed: ${(err as Error).message}`);
    }
  };

  private _toggleBacklinks = (fileRef: string): void => {
    const next = new Set(this.state.expandedBacklinks);
    if (next.has(fileRef)) next.delete(fileRef);
    else next.add(fileRef);
    this._safeSetState({ expandedBacklinks: next });
  };

  private _renderBacklinks(fileRef: string): React.ReactElement | null {
    const entry = this.state.backlinksByRef[fileRef];
    if (!entry) {
      // Not loaded yet (or no entry — same effective UI: nothing to show).
      if (this.state.backlinksLoading) {
        return (
          <div style={{ fontSize: 11, opacity: 0.6, marginTop: 2 }}>Loading backlinks...</div>
        );
      }
      return null;
    }
    const total = entry.visible.length + entry.hiddenCount;
    if (total === 0) {
      return (
        <div style={{ fontSize: 11, opacity: 0.6, marginTop: 2 }}>No incoming links found.</div>
      );
    }
    const expanded = this.state.expandedBacklinks.has(fileRef);
    return (
      <div style={{ fontSize: 11, marginTop: 2 }}>
        <a
          href="#"
          onClick={(e) => { e.preventDefault(); this._toggleBacklinks(fileRef); }}
          style={{ ...LINK_STYLE, cursor: 'pointer' }}
        >
          {expanded ? '▾' : '▸'} Linked from {entry.visible.length}
          {entry.hiddenCount > 0 ? ` (+${entry.hiddenCount} on sites you can't see)` : ''}
        </a>
        {expanded && entry.visible.length > 0 && (
          <ul style={{ margin: '4px 0 0 18px', padding: 0, listStyle: 'disc' }}>
            {entry.visible.slice(0, 50).map((s, i) => (
              <li key={i} style={{ marginBottom: 2 }}>
                <a href={s.url} {...EXTERNAL_LINK_ATTRS} style={LINK_STYLE}>
                  {s.title}
                </a>
                <span style={{ opacity: 0.7, marginLeft: 6 }}>
                  ({s.sourceKind}, {s.site})
                </span>
              </li>
            ))}
            {entry.visible.length > 50 && (
              <li style={{ opacity: 0.7 }}>...{entry.visible.length - 50} more</li>
            )}
          </ul>
        )}
      </div>
    );
  }

  private _export = async (type: 'exact' | 'stale' | 'samename'): Promise<void> => {
    this._safeSetState({ exporting: true });
    try {
      await this.props.service.downloadDuplicatesCsv(type);
    } catch (err) {
      this.props.onError(`Export ${type}: ${(err as Error).message}`);
    } finally {
      this._safeSetState({ exporting: false });
    }
  };

  private _addEntry = async (): Promise<void> => {
    const { addKind, addValue, addNote } = this.state;
    if (!addValue.trim() || !addNote.trim()) {
      this.props.onError('Allowlist: value and note are required');
      return;
    }
    this._safeSetState({ mutating: true });
    try {
      const input = {
        kind: addKind,
        note: addNote.trim(),
        ...(addKind === 'hash'
          ? { sha256: addValue.trim() }
          : { pattern: addValue.trim() }),
      };
      const res = await this.props.service.addDuplicatesAllowlistEntry(input);
      this._safeSetState({
        allowlist: res.allowlist,
        addValue: '',
        addNote: '',
        mutating: false,
      });
      // Refresh report to reflect new suppressions.
      void this._load();
    } catch (err) {
      this._safeSetState({ mutating: false });
      this.props.onError(`Allowlist add: ${(err as Error).message}`);
    }
  };

  private _removeEntry = async (
    kind: 'hash' | 'path' | 'name',
    payload: { sha256?: string; pattern?: string },
  ): Promise<void> => {
    this._safeSetState({ mutating: true });
    try {
      const res = await this.props.service.removeDuplicatesAllowlistEntry({ kind, ...payload });
      this._safeSetState({ allowlist: res.allowlist, mutating: false });
      void this._load();
    } catch (err) {
      this._safeSetState({ mutating: false });
      this.props.onError(`Allowlist remove: ${(err as Error).message}`);
    }
  };

  private _allowlistHash = (sha256: string, note: string): void => {
    void (async (): Promise<void> => {
      try {
        await this.props.service.addDuplicatesAllowlistEntry({ kind: 'hash', sha256, note });
        void this._load();
      } catch (err) {
        this.props.onError(`Allowlist hash: ${(err as Error).message}`);
      }
    })();
  };

  private _runBootstrap = async (): Promise<void> => {
    const {
      bootstrapSitePath,
      bootstrapLibraryTitle,
      bootstrapMaxFiles,
      bootstrapMaxVersions,
      bootstrapIncludeOther,
    } = this.state;
    const sitePath = bootstrapSitePath.trim();
    const libraryTitle = bootstrapLibraryTitle.trim();
    if (!sitePath || !libraryTitle) {
      this.props.onError('Bootstrap: sitePath and libraryTitle are required');
      return;
    }
    const maxFiles = Math.max(1, Math.min(Number.parseInt(bootstrapMaxFiles, 10) || 500, 5000));
    const maxVersionsPerFile = Math.max(1, Math.min(Number.parseInt(bootstrapMaxVersions, 10) || 20, 100));
    this._safeSetState({ bootstrapRunning: true, bootstrapActiveJob: undefined });
    try {
      const res = await this.props.service.bootstrapDuplicatesFromVersions({
        sitePath,
        libraryTitle,
        maxFiles,
        maxVersionsPerFile,
        includeOther: bootstrapIncludeOther,
      });
      if (!res.jobId) {
        // Empty-library 200 response (no files found).
        this._safeSetState({ bootstrapRunning: false });
        this.props.onError(res.info ?? 'Bootstrap: nothing to do');
        return;
      }
      // Kick off polling.
      this._startBootstrapPolling(res.jobId);
    } catch (err) {
      this._safeSetState({ bootstrapRunning: false });
      this.props.onError(`Bootstrap: ${(err as Error).message}`);
    }
  };

  private _startBootstrapPolling = (jobId: string): void => {
    if (this._bootstrapPollHandle !== undefined) {
      window.clearInterval(this._bootstrapPollHandle);
    }
    const tick = async (): Promise<void> => {
      if (!this._mounted) return;
      try {
        const job = await this.props.service.getDuplicatesBootstrapStatus(jobId);
        this._safeSetState({ bootstrapActiveJob: job });
        if (job.status === 'completed' || job.status === 'failed') {
          if (this._bootstrapPollHandle !== undefined) {
            window.clearInterval(this._bootstrapPollHandle);
            this._bootstrapPollHandle = undefined;
          }
          this._safeSetState({ bootstrapRunning: false });
          // Refresh report so new stale signals appear.
          void this._load();
        }
      } catch (err) {
        if (this._bootstrapPollHandle !== undefined) {
          window.clearInterval(this._bootstrapPollHandle);
          this._bootstrapPollHandle = undefined;
        }
        this._safeSetState({ bootstrapRunning: false });
        this.props.onError(`Bootstrap poll: ${(err as Error).message}`);
      }
    };
    void tick();
    this._bootstrapPollHandle = window.setInterval(() => { void tick(); }, 2000);
  };

  private _renderExact(): React.ReactElement {
    const { report, isAdmin } = this.state;
    if (!report) return <div />;
    if (report.exactGroups.length === 0) {
      return <Text variant="small">No exact duplicates detected.</Text>;
    }
    return (
      <Stack tokens={{ childrenGap: 8 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="small">
            <strong>{report.totals.exactGroups.toLocaleString()}</strong> group{report.totals.exactGroups === 1 ? '' : 's'}
            {report.truncated?.exactGroups ? ` (showing ${report.exactGroups.length})` : ''}
          </Text>
          {isAdmin && (
            <DefaultButton
              onClick={() => { void this._export('exact'); }}
              disabled={this.state.exporting}
              text="Export CSV"
            />
          )}
        </Stack>
        {report.exactGroups.map((g) => (
          <div
            key={g.sha256}
            style={{
              border: '1px solid rgba(127,127,127,0.3)',
              borderRadius: 2,
              padding: 8,
              background: 'rgba(127,127,127,0.04)',
            }}
          >
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text variant="small">
                <strong>{g.files.length} files</strong> · sha256 <code style={{ fontSize: 11 }}>{g.sha256.slice(0, 12)}…</code>
              </Text>
              {isAdmin && (
                <DefaultButton
                  onClick={() => this._allowlistHash(g.sha256, 'Dismissed from duplicates report')}
                  disabled={this.state.mutating}
                  title="Add this sha256 to the allowlist"
                  text="Dismiss group"
                />
              )}
            </Stack>
            <ul style={{ margin: '6px 0 0 18px', padding: 0, fontSize: 12 }}>
              {g.files.map((f) => (
                <li key={f.fileRef}>
                  <a href={toAbs(this.props.tenantOrigin, f.fileRef)} {...EXTERNAL_LINK_ATTRS} style={LINK_STYLE}>{f.fileName}</a>
                  {' · '}
                  <a
                    href={libraryUrlFromFileRef(this.props.tenantOrigin, f.sitePath, f.fileRef) ?? toAbs(this.props.tenantOrigin, f.sitePath)}
                    {...EXTERNAL_LINK_ATTRS}
                    style={LINK_STYLE_MONO}
                  >
                    {f.sitePath}/{f.library}
                  </a>
                  {' · '}
                  <span style={{ opacity: 0.7 }}>{fmtBytes(f.size)}</span>
                  {this._renderBacklinks(f.fileRef)}
                </li>
              ))}
            </ul>
          </div>
        ))}
      </Stack>
    );
  }

  private _renderStale(): React.ReactElement {
    const { report, isAdmin } = this.state;
    if (!report) return <div />;
    if (report.stalePairs.length === 0) {
      return <Text variant="small">No stale copies detected.</Text>;
    }
    return (
      <Stack tokens={{ childrenGap: 8 }}>
        <Stack horizontal horizontalAlign="space-between">
          <Text variant="small">
            <strong>{report.totals.staleFiles.toLocaleString()}</strong> stale file{report.totals.staleFiles === 1 ? '' : 's'}
            {report.truncated?.stalePairs ? ` (showing ${report.stalePairs.length} pairs)` : ''}
          </Text>
          {isAdmin && (
            <DefaultButton
              onClick={() => { void this._export('stale'); }}
              disabled={this.state.exporting}
              text="Export CSV"
            />
          )}
        </Stack>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ borderBottom: '1px solid rgba(127,127,127,0.4)', textAlign: 'left' }}>
              <th style={{ padding: 6 }}>Stale file</th>
              <th style={{ padding: 6 }}>Authoritative file</th>
              <th style={{ padding: 6 }}>Diverged at</th>
            </tr>
          </thead>
          <tbody>
            {report.stalePairs.map((p, i) => (
              <tr key={i} style={{ borderBottom: '1px solid rgba(127,127,127,0.2)' }}>
                <td style={{ padding: 6 }}>
                  <a href={toAbs(this.props.tenantOrigin, p.staleFileRef)} {...EXTERNAL_LINK_ATTRS} style={LINK_STYLE}>{p.staleFileName}</a>
                  <div style={{ fontSize: 11 }}>
                    <a
                      href={libraryUrlFromFileRef(this.props.tenantOrigin, p.staleSitePath, p.staleFileRef) ?? toAbs(this.props.tenantOrigin, p.staleSitePath)}
                      {...EXTERNAL_LINK_ATTRS}
                      style={{ ...LINK_STYLE_MONO, fontSize: 11 }}
                    >
                      {p.staleSitePath}/{p.staleLibrary}
                    </a>
                  </div>
                  {this._renderBacklinks(p.staleFileRef)}
                </td>
                <td style={{ padding: 6 }}>
                  <a href={toAbs(this.props.tenantOrigin, p.authoritativeFileRef)} {...EXTERNAL_LINK_ATTRS} style={LINK_STYLE}>{p.authoritativeFileName}</a>
                  <div style={{ fontSize: 11 }}>
                    <a
                      href={libraryUrlFromFileRef(this.props.tenantOrigin, p.authoritativeSitePath, p.authoritativeFileRef) ?? toAbs(this.props.tenantOrigin, p.authoritativeSitePath)}
                      {...EXTERNAL_LINK_ATTRS}
                      style={{ ...LINK_STYLE_MONO, fontSize: 11 }}
                    >
                      {p.authoritativeSitePath}/{p.authoritativeLibrary}
                    </a>
                  </div>
                  {this._renderBacklinks(p.authoritativeFileRef)}
                </td>
                <td style={{ padding: 6, fontSize: 11 }}>{p.divergedAt.slice(0, 10)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Stack>
    );
  }

  private _renderDiverged(): React.ReactElement {
    const { report } = this.state;
    if (!report) return <div />;
    if (report.divergedPairs.length === 0) {
      return (
        <Text variant="small">
          No diverged pairs detected. Diverged means two files share a common ancestor in
          version history but both have edited away from it independently.
        </Text>
      );
    }
    return (
      <Stack tokens={{ childrenGap: 8 }}>
        <Text variant="small">
          <strong>{report.totals.divergedPairs.toLocaleString()}</strong> pair
          {report.totals.divergedPairs === 1 ? '' : 's'} where two files share a common
          ancestor in version history but currents differ
          {report.truncated?.divergedPairs ? ` (showing ${report.divergedPairs.length})` : ''}
        </Text>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ borderBottom: '1px solid rgba(127,127,127,0.4)', textAlign: 'left' }}>
              <th style={{ padding: 6 }}>File A</th>
              <th style={{ padding: 6 }}>File B</th>
              <th style={{ padding: 6, width: 110 }}>Ancestor at</th>
            </tr>
          </thead>
          <tbody>
            {report.divergedPairs.map((p, i) => (
              <tr key={i} style={{ borderBottom: '1px solid rgba(127,127,127,0.2)' }}>
                <td style={{ padding: 6 }}>
                  <a href={toAbs(this.props.tenantOrigin, p.aFileRef)} {...EXTERNAL_LINK_ATTRS} style={LINK_STYLE}>{p.aFileName}</a>
                  <div style={{ fontSize: 11 }}>
                    <a
                      href={libraryUrlFromFileRef(this.props.tenantOrigin, p.aSitePath, p.aFileRef) ?? toAbs(this.props.tenantOrigin, p.aSitePath)}
                      {...EXTERNAL_LINK_ATTRS}
                      style={{ ...LINK_STYLE_MONO, fontSize: 11 }}
                    >
                      {p.aSitePath}/{p.aLibrary}
                    </a>
                  </div>
                  {this._renderBacklinks(p.aFileRef)}
                </td>
                <td style={{ padding: 6 }}>
                  <a href={toAbs(this.props.tenantOrigin, p.bFileRef)} {...EXTERNAL_LINK_ATTRS} style={LINK_STYLE}>{p.bFileName}</a>
                  <div style={{ fontSize: 11 }}>
                    <a
                      href={libraryUrlFromFileRef(this.props.tenantOrigin, p.bSitePath, p.bFileRef) ?? toAbs(this.props.tenantOrigin, p.bSitePath)}
                      {...EXTERNAL_LINK_ATTRS}
                      style={{ ...LINK_STYLE_MONO, fontSize: 11 }}
                    >
                      {p.bSitePath}/{p.bLibrary}
                    </a>
                  </div>
                  {this._renderBacklinks(p.bFileRef)}
                </td>
                <td style={{ padding: 6, fontSize: 11 }}>{p.ancestorObservedAt.slice(0, 10)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Stack>
    );
  }

  private _renderNearDuplicate(): React.ReactElement {
    const { report } = this.state;
    if (!report) return <div />;
    if (report.nearDuplicatePairs.length === 0) {
      return (
        <Text variant="small">
          No near-duplicates detected. Near-duplicates have visible text similar enough that a
          64-bit SimHash matches within a small Hamming distance — typo fixes, sentence
          reorderings, or polish-only edits.
        </Text>
      );
    }
    return (
      <Stack tokens={{ childrenGap: 8 }}>
        <Text variant="small">
          <strong>{report.totals.nearDuplicatePairs.toLocaleString()}</strong> pair
          {report.totals.nearDuplicatePairs === 1 ? '' : 's'} within SimHash threshold
          {report.truncated?.nearDuplicatePairs ? ` (showing ${report.nearDuplicatePairs.length})` : ''}
        </Text>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ borderBottom: '1px solid rgba(127,127,127,0.4)', textAlign: 'left' }}>
              <th style={{ padding: 6 }}>File A</th>
              <th style={{ padding: 6 }}>File B</th>
              <th style={{ padding: 6, width: 110 }}>Hamming dist.</th>
            </tr>
          </thead>
          <tbody>
            {report.nearDuplicatePairs.map((p, i) => (
              <tr key={i} style={{ borderBottom: '1px solid rgba(127,127,127,0.2)' }}>
                <td style={{ padding: 6 }}>
                  <a href={toAbs(this.props.tenantOrigin, p.aFileRef)} {...EXTERNAL_LINK_ATTRS} style={LINK_STYLE}>{p.aFileName}</a>
                  <div style={{ fontSize: 11 }}>
                    <a
                      href={libraryUrlFromFileRef(this.props.tenantOrigin, p.aSitePath, p.aFileRef) ?? toAbs(this.props.tenantOrigin, p.aSitePath)}
                      {...EXTERNAL_LINK_ATTRS}
                      style={{ ...LINK_STYLE_MONO, fontSize: 11 }}
                    >
                      {p.aSitePath}/{p.aLibrary}
                    </a>
                  </div>
                  {this._renderBacklinks(p.aFileRef)}
                </td>
                <td style={{ padding: 6 }}>
                  <a href={toAbs(this.props.tenantOrigin, p.bFileRef)} {...EXTERNAL_LINK_ATTRS} style={LINK_STYLE}>{p.bFileName}</a>
                  <div style={{ fontSize: 11 }}>
                    <a
                      href={libraryUrlFromFileRef(this.props.tenantOrigin, p.bSitePath, p.bFileRef) ?? toAbs(this.props.tenantOrigin, p.bSitePath)}
                      {...EXTERNAL_LINK_ATTRS}
                      style={{ ...LINK_STYLE_MONO, fontSize: 11 }}
                    >
                      {p.bSitePath}/{p.bLibrary}
                    </a>
                  </div>
                  {this._renderBacklinks(p.bFileRef)}
                </td>
                <td style={{ padding: 6, fontSize: 11 }}>{p.hammingDistance} / 64</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Stack>
    );
  }

  private _renderSameName(): React.ReactElement {
    const { report, isAdmin } = this.state;
    if (!report) return <div />;
    if (report.sameNamePairs.length === 0) {
      return <Text variant="small">No same-name warnings.</Text>;
    }
    return (
      <Stack tokens={{ childrenGap: 8 }}>
        <Stack horizontal horizontalAlign="space-between">
          <Text variant="small">
            <strong>{report.totals.sameNamePairs.toLocaleString()}</strong> pair{report.totals.sameNamePairs === 1 ? '' : 's'}
            {report.truncated?.sameNamePairs ? ` (showing ${report.sameNamePairs.length})` : ''}
          </Text>
          {isAdmin && (
            <DefaultButton
              onClick={() => { void this._export('samename'); }}
              disabled={this.state.exporting}
              text="Export CSV"
            />
          )}
        </Stack>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ borderBottom: '1px solid rgba(127,127,127,0.4)', textAlign: 'left' }}>
              <th style={{ padding: 6 }}>File A</th>
              <th style={{ padding: 6 }}>File B</th>
              <th style={{ padding: 6, width: 90 }}>Size match</th>
              <th style={{ padding: 6, width: 120 }}></th>
            </tr>
          </thead>
          <tbody>
            {report.sameNamePairs.map((p, i) => (
              <tr key={i} style={{ borderBottom: '1px solid rgba(127,127,127,0.2)' }}>
                <td style={{ padding: 6 }}>
                  <a href={toAbs(this.props.tenantOrigin, p.aFileRef)} {...EXTERNAL_LINK_ATTRS} style={LINK_STYLE}>{p.aFileName}</a>
                  <div style={{ fontSize: 11 }}>
                    <a
                      href={libraryUrlFromFileRef(this.props.tenantOrigin, p.aSitePath, p.aFileRef) ?? toAbs(this.props.tenantOrigin, p.aSitePath)}
                      {...EXTERNAL_LINK_ATTRS}
                      style={{ ...LINK_STYLE_MONO, fontSize: 11 }}
                    >
                      {p.aSitePath}
                    </a>
                    {' · '}
                    <span style={{ opacity: 0.7 }}>{fmtBytes(p.aSize)}</span>
                  </div>
                  {this._renderBacklinks(p.aFileRef)}
                </td>
                <td style={{ padding: 6 }}>
                  <a href={toAbs(this.props.tenantOrigin, p.bFileRef)} {...EXTERNAL_LINK_ATTRS} style={LINK_STYLE}>{p.bFileName}</a>
                  <div style={{ fontSize: 11 }}>
                    <a
                      href={libraryUrlFromFileRef(this.props.tenantOrigin, p.bSitePath, p.bFileRef) ?? toAbs(this.props.tenantOrigin, p.bSitePath)}
                      {...EXTERNAL_LINK_ATTRS}
                      style={{ ...LINK_STYLE_MONO, fontSize: 11 }}
                    >
                      {p.bSitePath}
                    </a>
                    {' · '}
                    <span style={{ opacity: 0.7 }}>{fmtBytes(p.bSize)}</span>
                  </div>
                  {this._renderBacklinks(p.bFileRef)}
                </td>
                <td style={{ padding: 6 }}>
                  {p.sameSize ? (
                    <span style={{ color: 'var(--warningIcon, var(--orangeLight, #d47800))' }}>⚠ Yes</span>
                  ) : (
                    <span style={{ opacity: 0.7 }}>No</span>
                  )}
                </td>
                <td style={{ padding: 6 }}>
                  {isAdmin && (
                    <DefaultButton
                      onClick={() => { void this._addEntryInline('name', p.aFileName, `Allow multi-site ${p.aFileName}`); }}
                      disabled={this.state.mutating}
                      text="Silence name"
                    />
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Stack>
    );
  }

  private _addEntryInline = async (
    kind: 'hash' | 'path' | 'name',
    value: string,
    note: string,
  ): Promise<void> => {
    this._safeSetState({ mutating: true });
    try {
      const input = {
        kind,
        note,
        ...(kind === 'hash' ? { sha256: value } : { pattern: value }),
      };
      const res = await this.props.service.addDuplicatesAllowlistEntry(input);
      this._safeSetState({ allowlist: res.allowlist, mutating: false });
      void this._load();
    } catch (err) {
      this._safeSetState({ mutating: false });
      this.props.onError(`Allowlist add: ${(err as Error).message}`);
    }
  };

  private _renderAllowlist(): React.ReactElement {
    const { allowlist, isAdmin, addKind, addValue, addNote, mutating } = this.state;
    if (!allowlist) return <div />;
    const section = (
      title: string,
      kind: 'hash' | 'path' | 'name',
      entries: Array<{ note: string; addedBy: string; addedAt: string; sha256?: string; pattern?: string }>,
    ): React.ReactElement => (
      <div style={{ marginBottom: 14 }}>
        <Text variant="smallPlus" style={{ fontWeight: 600 }}>{title} ({entries.length})</Text>
        {entries.length === 0 ? (
          <div style={{ opacity: 0.7, fontSize: 12, marginTop: 4 }}>(none)</div>
        ) : (
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12, marginTop: 4 }}>
            <tbody>
              {entries.map((e, i) => (
                <tr key={i} style={{ borderBottom: '1px solid rgba(127,127,127,0.2)' }}>
                  <td style={{ padding: 4, fontFamily: 'Consolas, monospace' }}>
                    {kind === 'hash' ? e.sha256?.slice(0, 20) + '…' : e.pattern}
                  </td>
                  <td style={{ padding: 4 }}>{e.note}</td>
                  <td style={{ padding: 4, opacity: 0.7 }}>{e.addedBy}</td>
                  <td style={{ padding: 4, opacity: 0.7 }}>{e.addedAt.slice(0, 10)}</td>
                  <td style={{ padding: 4 }}>
                    {isAdmin && (
                      <DefaultButton
                        onClick={() => {
                          void this._removeEntry(kind, kind === 'hash'
                            ? { sha256: e.sha256 }
                            : { pattern: e.pattern });
                        }}
                        disabled={mutating}
                        text="Remove"
                      />
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    );
    return (
      <Stack tokens={{ childrenGap: 10 }}>
        {isAdmin && (
          <div
            style={{
              border: '1px solid rgba(127,127,127,0.3)',
              padding: 10,
              borderRadius: 2,
              background: 'rgba(127,127,127,0.04)',
            }}
          >
            <Text variant="smallPlus" style={{ fontWeight: 600, marginBottom: 6, display: 'block' }}>
              Add allowlist entry
            </Text>
            <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="end" wrap>
              <select
                value={addKind}
                onChange={(e) => this._safeSetState({ addKind: e.target.value as 'hash' | 'path' | 'name' })}
                style={{ padding: 6 }}
              >
                <option value="hash">hash (sha256)</option>
                <option value="path">path (glob)</option>
                <option value="name">name (glob)</option>
              </select>
              <TextField
                placeholder={addKind === 'hash' ? 'sha256 hex' : '/sites/*/Templates/**'}
                value={addValue}
                onChange={(_, v) => this._safeSetState({ addValue: v ?? '' })}
                styles={{ root: { minWidth: 280 } }}
              />
              <TextField
                placeholder="note (required)"
                value={addNote}
                onChange={(_, v) => this._safeSetState({ addNote: v ?? '' })}
                styles={{ root: { minWidth: 220 } }}
              />
              <PrimaryButton
                disabled={mutating || !addValue.trim() || !addNote.trim()}
                onClick={() => { void this._addEntry(); }}
                text="Add"
              />
            </Stack>
          </div>
        )}
        {section('Hash allowlist', 'hash', allowlist.hashAllowlist)}
        {section('Path allowlist', 'path', allowlist.pathAllowlist)}
        {section('Name allowlist', 'name', allowlist.nameAllowlist)}
      </Stack>
    );
  }

  private _renderBootstrap(): React.ReactElement {
    const {
      bootstrapSitePath,
      bootstrapLibraryTitle,
      bootstrapMaxFiles,
      bootstrapMaxVersions,
      bootstrapIncludeOther,
      bootstrapRunning,
      bootstrapActiveJob,
      isAdmin,
    } = this.state;
    return (
      <Stack tokens={{ childrenGap: 10 }}>
        <MessageBar messageBarType={MessageBarType.info}>
          <strong>Version-history bootstrap.</strong> Reads the complete version
          history of every file in the selected library and backfills
          previous-hash entries into the duplicate-detection index. This is a
          one-shot operation per library &mdash; run it once to seed stale-copy
          detection; after that, weekly scans keep the index fresh on their own.
        </MessageBar>
        {!isAdmin && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Admin-only. Your account can&apos;t trigger a bootstrap.
          </MessageBar>
        )}
        <div
          style={{
            border: '1px solid rgba(127,127,127,0.3)',
            borderRadius: 2,
            padding: 12,
            background: 'rgba(127,127,127,0.04)',
          }}
        >
          <Stack tokens={{ childrenGap: 8 }}>
            <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
              <TextField
                label="Site path"
                placeholder="/sites/hub"
                value={bootstrapSitePath}
                onChange={(_, v) => this._safeSetState({ bootstrapSitePath: v ?? '' })}
                disabled={bootstrapRunning}
                styles={{ root: { minWidth: 260 } }}
              />
              <TextField
                label="Library title"
                placeholder="Documents"
                value={bootstrapLibraryTitle}
                onChange={(_, v) => this._safeSetState({ bootstrapLibraryTitle: v ?? '' })}
                disabled={bootstrapRunning}
                styles={{ root: { minWidth: 200 } }}
              />
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end" wrap>
              <TextField
                label="Max files"
                description="1–5000 (default 500)"
                value={bootstrapMaxFiles}
                onChange={(_, v) => this._safeSetState({ bootstrapMaxFiles: v ?? '' })}
                disabled={bootstrapRunning}
                styles={{ root: { width: 140 } }}
              />
              <TextField
                label="Max versions per file"
                description="1–100 (default 20)"
                value={bootstrapMaxVersions}
                onChange={(_, v) => this._safeSetState({ bootstrapMaxVersions: v ?? '' })}
                disabled={bootstrapRunning}
                styles={{ root: { width: 180 } }}
              />
              <Checkbox
                label="Include non-Office files"
                checked={bootstrapIncludeOther}
                onChange={(_, checked) => this._safeSetState({ bootstrapIncludeOther: !!checked })}
                disabled={bootstrapRunning}
              />
              <PrimaryButton
                disabled={!isAdmin || bootstrapRunning || !bootstrapSitePath.trim() || !bootstrapLibraryTitle.trim()}
                onClick={() => { void this._runBootstrap(); }}
                title={
                  !isAdmin
                    ? 'Admin only'
                    : !bootstrapSitePath.trim()
                      ? 'Enter a site path (e.g. /sites/hub) to enable'
                      : !bootstrapLibraryTitle.trim()
                        ? 'Enter a library title (e.g. Documents) to enable'
                        : bootstrapRunning
                          ? 'A bootstrap is already running'
                          : 'Run version-history bootstrap on this library'
                }
                text={bootstrapRunning ? 'Running...' : 'Run bootstrap'}
              />
            </Stack>
          </Stack>
        </div>
        {bootstrapActiveJob && (
          <div
            style={{
              border: '1px solid rgba(127,127,127,0.3)',
              borderRadius: 2,
              padding: 10,
              background: 'rgba(127,127,127,0.04)',
            }}
          >
            <Text variant="smallPlus" style={{ fontWeight: 600, display: 'block', marginBottom: 4 }}>
              Bootstrap job {bootstrapActiveJob.jobId.slice(-12)} · {bootstrapActiveJob.status}
              {' · '}
              {bootstrapActiveJob.sitePath}/{bootstrapActiveJob.libraryTitle}
            </Text>
            <Text variant="small" style={{ display: 'block' }}>
              <strong>{bootstrapActiveJob.filesCompleted.toLocaleString()}</strong>
              {' / '}
              <strong>{bootstrapActiveJob.filesTotal.toLocaleString()}</strong> files
              {' · '}
              <strong>{bootstrapActiveJob.versionsProcessed.toLocaleString()}</strong> versions hashed
              {bootstrapActiveJob.errorCount > 0 ? ` · ${bootstrapActiveJob.errorCount} errors` : ''}
            </Text>
            {bootstrapActiveJob.filesTotal > 0 && (
              <div
                style={{
                  marginTop: 6,
                  height: 6,
                  background: 'rgba(127,127,127,0.2)',
                  borderRadius: 3,
                  overflow: 'hidden',
                }}
              >
                <div
                  style={{
                    height: '100%',
                    width: `${Math.min(100, (bootstrapActiveJob.filesCompleted / bootstrapActiveJob.filesTotal) * 100)}%`,
                    background: 'var(--communicationPrimary, #3a96dd)',
                    transition: 'width 0.3s ease',
                  }}
                />
              </div>
            )}
            {bootstrapActiveJob.currentFile && bootstrapActiveJob.status === 'running' && (
              <Text variant="small" style={{ display: 'block', marginTop: 4, opacity: 0.8, fontFamily: 'Consolas, monospace' }}>
                {bootstrapActiveJob.currentFile}
              </Text>
            )}
            {bootstrapActiveJob.recentErrors.length > 0 && (
              <details style={{ marginTop: 6 }}>
                <summary style={{ cursor: 'pointer', fontSize: 12 }}>Show errors</summary>
                <ul style={{ fontSize: 11, fontFamily: 'Consolas, monospace', margin: '4px 0 0 18px' }}>
                  {bootstrapActiveJob.recentErrors.slice(0, 50).map((e, i) => (
                    <li key={i}>{e}</li>
                  ))}
                </ul>
              </details>
            )}
          </div>
        )}
      </Stack>
    );
  }

  public render(): React.ReactElement {
    const { loading, report, isAdmin } = this.state;
    if (loading) {
      return (
        <Stack horizontalAlign="center" verticalAlign="center" style={{ padding: 40 }}>
          <Spinner size={SpinnerSize.large} label="Loading duplicates..." />
        </Stack>
      );
    }
    if (!report || !report.indexBuiltAt) {
      return (
        <MessageBar messageBarType={MessageBarType.info}>
          No hash index yet. Run a document scan to populate it, then return here.
        </MessageBar>
      );
    }
    return (
      <Stack tokens={{ childrenGap: 10 }} style={{ padding: 10 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="medium">
            Hash index built: <strong>{report.indexBuiltAt.slice(0, 19).replace('T', ' ')} UTC</strong>
            {' · '}
            {report.totals.exactGroups} exact · {report.totals.staleFiles} stale · {report.totals.divergedPairs} diverged · {report.totals.nearDuplicatePairs} near-duplicate · {report.totals.sameNamePairs} same-name
          </Text>
          <DefaultButton onClick={() => { void this._load(); }} iconProps={{ iconName: 'Refresh' }} text="Refresh" />
        </Stack>
        {!isAdmin && (
          <MessageBar messageBarType={MessageBarType.info}>
            Showing duplicates on sites you can read. Admin-only actions are hidden.
          </MessageBar>
        )}
        <Pivot>
          <PivotItem headerText={`Exact (${report.totals.exactGroups})`} itemKey="exact">
            <div style={{ padding: 10 }}>{this._renderExact()}</div>
          </PivotItem>
          <PivotItem headerText={`Stale (${report.totals.staleFiles})`} itemKey="stale">
            <div style={{ padding: 10 }}>{this._renderStale()}</div>
          </PivotItem>
          <PivotItem headerText={`Diverged (${report.totals.divergedPairs})`} itemKey="diverged">
            <div style={{ padding: 10 }}>{this._renderDiverged()}</div>
          </PivotItem>
          <PivotItem headerText={`Near-duplicate (${report.totals.nearDuplicatePairs})`} itemKey="nearduplicate">
            <div style={{ padding: 10 }}>{this._renderNearDuplicate()}</div>
          </PivotItem>
          <PivotItem headerText={`Same-name (${report.totals.sameNamePairs})`} itemKey="samename">
            <div style={{ padding: 10 }}>{this._renderSameName()}</div>
          </PivotItem>
          <PivotItem headerText="Allowlist" itemKey="allowlist">
            <div style={{ padding: 10 }}>{this._renderAllowlist()}</div>
          </PivotItem>
          <PivotItem headerText="Bootstrap" itemKey="bootstrap">
            <div style={{ padding: 10 }}>{this._renderBootstrap()}</div>
          </PivotItem>
        </Pivot>
      </Stack>
    );
  }
}
