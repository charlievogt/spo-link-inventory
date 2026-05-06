import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import {
  LinkInventoryService,
  type IOrphanFile,
  type IOrphanRecycleResponse,
  type IOrphansReport,
} from '../services/LinkInventoryService';

interface IOrphansTabProps {
  service: LinkInventoryService;
  onError: (message: string) => void;
}

interface IOrphansTabState {
  /** Initial load (or a manual refresh). */
  loading: boolean;
  /** True after the user clicked Recycle, until the response comes back. */
  recycling: boolean;
  report?: IOrphansReport;
  /** Set of canonical file URLs the user has selected for recycle. */
  selected: Set<string>;
  /** Per-site collapse/expand state (sites with no orphans default collapsed). */
  collapsed: Set<string>;
  /** Last recycle run's response, if any. Drives the result banner. */
  lastRunResult?: IOrphanRecycleResponse;
  /** When true, the user has explicitly acknowledged a stale-index hard block. */
  acknowledgedStaleIndex: boolean;
  /** Server-returned 409 stale message — distinct from the soft warn banner. */
  staleBlockMessage?: string;
}

const EXTERNAL_LINK_ATTRS = {
  target: '_blank',
  rel: 'noopener noreferrer',
  'data-interception': 'off',
} as const;

const LINK_STYLE: React.CSSProperties = {
  color: 'var(--communicationPrimary, #3a96dd)',
  textDecoration: 'none',
};
const MONO_STYLE: React.CSSProperties = {
  fontFamily: 'Consolas, monospace',
  fontSize: '0.85em',
  color: 'var(--neutralSecondary, #605e5c)',
  wordBreak: 'break-all',
};

function fmtBytes(b: number): string {
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / 1024 / 1024).toFixed(1)} MB`;
}

function fmtAge(min: number): string {
  if (min < 60) return `${min} min`;
  if (min < 60 * 24) return `${Math.round(min / 60)} h`;
  return `${Math.round(min / (60 * 24))} d`;
}

export class OrphansTab extends React.Component<IOrphansTabProps, IOrphansTabState> {
  private _mounted = false;

  public constructor(props: IOrphansTabProps) {
    super(props);
    this.state = {
      loading: true,
      recycling: false,
      selected: new Set<string>(),
      collapsed: new Set<string>(),
      acknowledgedStaleIndex: false,
    };
  }

  public componentDidMount(): void {
    this._mounted = true;
    void this._load();
  }

  public componentWillUnmount(): void {
    this._mounted = false;
  }

  private _safeSetState = <K extends keyof IOrphansTabState>(
    patch: Pick<IOrphansTabState, K>,
  ): void => {
    if (!this._mounted) return;
    this.setState(patch);
  };

  private _load = async (overrides?: { acknowledgeStale?: boolean }): Promise<void> => {
    // setState is async — callers that just flipped acknowledgedStaleIndex
    // need to pass it explicitly here, otherwise _load reads the stale value.
    const acknowledgedStaleIndex =
      overrides?.acknowledgeStale ?? this.state.acknowledgedStaleIndex;
    this._safeSetState({ loading: true, lastRunResult: undefined, staleBlockMessage: undefined });
    try {
      const report = await this.props.service.getOrphansReport({
        acknowledgedStaleIndex,
      });
      // Auto-collapse sites with zero orphans so the screen isn't dominated
      // by clean sites; the user only sees cluttered sites by default.
      const collapsed = new Set<string>();
      for (const site of report.sites) {
        if (site.orphans.length === 0) collapsed.add(site.sitePath);
      }
      this._safeSetState({
        loading: false,
        report,
        selected: new Set<string>(),
        collapsed,
      });
    } catch (err) {
      const message = (err as Error).message;
      // Surface the 409 stale-block message inline (caller can override),
      // not as a top-level error banner.
      if (/Backlinks index is .+ min old/i.test(message)) {
        this._safeSetState({
          loading: false,
          staleBlockMessage: message,
        });
      } else if (/no persistent backlinks index/i.test(message)) {
        this._safeSetState({
          loading: false,
          staleBlockMessage:
            'No backlinks index yet. Run a tenant page-scan first (Link Inventory tab → Run scan).',
        });
      } else {
        this.props.onError(message);
        this._safeSetState({ loading: false });
      }
    }
  };

  private _toggleSite = (sitePath: string): void => {
    const next = new Set(this.state.collapsed);
    if (next.has(sitePath)) next.delete(sitePath);
    else next.add(sitePath);
    this._safeSetState({ collapsed: next });
  };

  private _toggleFile = (serverRelativeUrl: string): void => {
    const next = new Set(this.state.selected);
    if (next.has(serverRelativeUrl)) next.delete(serverRelativeUrl);
    else next.add(serverRelativeUrl);
    this._safeSetState({ selected: next });
  };

  private _selectAllInSite = (sitePath: string, orphans: IOrphanFile[]): void => {
    const next = new Set(this.state.selected);
    const allSelected = orphans.every((o) => next.has(o.serverRelativeUrl));
    for (const o of orphans) {
      if (allSelected) next.delete(o.serverRelativeUrl);
      else next.add(o.serverRelativeUrl);
    }
    this._safeSetState({ selected: next });
  };

  private _clearSelection = (): void => {
    this._safeSetState({ selected: new Set<string>() });
  };

  /** Build the file list payload from the current selection. */
  private _selectedFilesPayload(): Array<{ sitePath: string; serverRelativeUrl: string }> {
    if (!this.state.report) return [];
    const out: Array<{ sitePath: string; serverRelativeUrl: string }> = [];
    for (const site of this.state.report.sites) {
      for (const o of site.orphans) {
        if (this.state.selected.has(o.serverRelativeUrl)) {
          out.push({ sitePath: site.sitePath, serverRelativeUrl: o.serverRelativeUrl });
        }
      }
    }
    return out;
  }

  private _runRecycle = async (dryRun: boolean): Promise<void> => {
    const files = this._selectedFilesPayload();
    if (files.length === 0) return;
    this._safeSetState({ recycling: true, lastRunResult: undefined });
    try {
      const res = await this.props.service.recycleOrphans({
        files,
        dryRun,
        confirmReportGeneratedAt: this.state.report?.generatedAt,
      });
      this._safeSetState({ recycling: false, lastRunResult: res });
      // After a real recycle, refresh the report so the recycled rows
      // disappear and the user can see the new clean state. Skip on
      // dry-run — the data didn't change.
      if (!dryRun) {
        void this._load();
      }
    } catch (err) {
      this.props.onError((err as Error).message);
      this._safeSetState({ recycling: false });
    }
  };

  public render(): React.ReactElement<IOrphansTabProps> {
    const { loading, recycling, report, selected, collapsed, lastRunResult, staleBlockMessage } = this.state;

    if (loading) {
      return (
        <Stack tokens={{ childrenGap: 12 }} styles={{ root: { padding: 16 } }}>
          <Spinner size={SpinnerSize.large} label="Loading orphan-asset report…" />
        </Stack>
      );
    }

    if (staleBlockMessage) {
      return (
        <Stack tokens={{ childrenGap: 12 }} styles={{ root: { padding: 16 } }}>
          <MessageBar messageBarType={MessageBarType.warning}>
            {staleBlockMessage}
          </MessageBar>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="Run anyway (override stale-index check)"
              onClick={(): void => {
                this._safeSetState({ acknowledgedStaleIndex: true });
                void this._load({ acknowledgeStale: true });
              }}
            />
            <DefaultButton text="Retry" onClick={(): void => { void this._load(); }} />
          </Stack>
        </Stack>
      );
    }

    if (!report) return <></>;

    const selectionCount = selected.size;
    const recycleAllowed = selectionCount > 0 && !recycling;

    return (
      <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: 16 } }}>
        {/* Index freshness note. The hard block is server-enforced; this
            is the soft warn. */}
        {report.warning && (
          <MessageBar messageBarType={MessageBarType.warning}>
            {report.warning}
          </MessageBar>
        )}

        {/* Header / totals */}
        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="xLarge">Orphan assets in SiteAssets/SitePages</Text>
          <Text variant="medium">
            {report.totals.orphans} orphan file
            {report.totals.orphans === 1 ? '' : 's'} across {report.totals.sitesScanned} site
            {report.totals.sitesScanned === 1 ? '' : 's'} • scanned {report.totals.filesScanned}{' '}
            file{report.totals.filesScanned === 1 ? '' : 's'} • backlinks index age:{' '}
            {fmtAge(report.indexAge.ageMinutes)}
          </Text>
        </Stack>

        {/* Action bar */}
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" wrap>
          <DefaultButton
            text="Refresh"
            onClick={(): void => { void this._load(); }}
            disabled={recycling}
            iconProps={{ iconName: 'Refresh' }}
          />
          <DefaultButton
            text={`Preview recycle (${selectionCount})`}
            onClick={(): void => { void this._runRecycle(true); }}
            disabled={!recycleAllowed}
          />
          <PrimaryButton
            text={`Recycle selected (${selectionCount})`}
            onClick={(): void => { void this._runRecycle(false); }}
            disabled={!recycleAllowed}
            iconProps={{ iconName: 'Delete' }}
          />
          {selectionCount > 0 && (
            <DefaultButton text="Clear selection" onClick={this._clearSelection} disabled={recycling} />
          )}
          {recycling && <Spinner size={SpinnerSize.small} label="Working…" />}
        </Stack>

        {/* Last run result */}
        {lastRunResult && (
          <MessageBar
            messageBarType={
              lastRunResult.summary.error > 0 ? MessageBarType.warning : MessageBarType.success
            }
            onDismiss={(): void => this._safeSetState({ lastRunResult: undefined })}
            dismissButtonAriaLabel="Close"
          >
            {lastRunResult.dryRun
              ? `Dry-run: would recycle ${lastRunResult.summary.preview} file${
                  lastRunResult.summary.preview === 1 ? '' : 's'
                }.`
              : `Run ${lastRunResult.runId.slice(0, 8)}: recycled ${lastRunResult.summary.recycled}, ` +
                `errors ${lastRunResult.summary.error}, ` +
                `not found ${lastRunResult.summary.notFound}, ` +
                `forbidden ${lastRunResult.summary.forbidden}` +
                (lastRunResult.summary.droppedSites.length > 0
                  ? `, dropped ${lastRunResult.summary.droppedSites.length} site(s) for write perms`
                  : '')}
          </MessageBar>
        )}

        {/* Per-site grouped list */}
        {report.sites.length === 0 && (
          <Text>No sites in scope. Run a tenant page-scan first.</Text>
        )}

        {report.sites.map((site) => {
          const isCollapsed = collapsed.has(site.sitePath);
          const orphanCount = site.orphans.length;
          const allSelected = orphanCount > 0 && site.orphans.every((o) => selected.has(o.serverRelativeUrl));
          return (
            <Stack
              key={site.sitePath}
              tokens={{ childrenGap: 4 }}
              styles={{
                root: {
                  border: '1px solid var(--neutralLight, #edebe9)',
                  borderRadius: 4,
                  padding: 12,
                  background:
                    orphanCount === 0
                      ? 'var(--neutralLighterAlt, #faf9f8)'
                      : 'var(--white, #fff)',
                },
              }}
            >
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <DefaultButton
                  iconProps={{ iconName: isCollapsed ? 'ChevronRight' : 'ChevronDown' }}
                  ariaLabel={isCollapsed ? 'Expand' : 'Collapse'}
                  onClick={(): void => this._toggleSite(site.sitePath)}
                  styles={{ root: { minWidth: 32, padding: 0 } }}
                />
                <Stack grow>
                  <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                    {site.sitePath}{' '}
                    <Text variant="small" styles={{ root: { color: 'var(--neutralSecondary, #605e5c)' } }}>
                      ({orphanCount} orphan{orphanCount === 1 ? '' : 's'} of {site.filesScanned} scanned)
                    </Text>
                  </Text>
                  {site.error && (
                    <Text variant="small" styles={{ root: { color: 'var(--errorText, #a4262c)' } }}>
                      Error: {site.error}
                    </Text>
                  )}
                </Stack>
                {orphanCount > 0 && (
                  <DefaultButton
                    text={allSelected ? 'Deselect all' : 'Select all'}
                    onClick={(): void => this._selectAllInSite(site.sitePath, site.orphans)}
                    disabled={recycling}
                  />
                )}
              </Stack>

              {!isCollapsed && orphanCount > 0 && (
                <Stack tokens={{ childrenGap: 4 }} styles={{ root: { paddingLeft: 40, marginTop: 4 } }}>
                  {site.orphans.map((o) => (
                    <Stack
                      key={o.serverRelativeUrl}
                      horizontal
                      tokens={{ childrenGap: 8 }}
                      verticalAlign="start"
                      styles={{ root: { borderTop: '1px solid var(--neutralLighter, #f3f2f1)', paddingTop: 6, paddingBottom: 6 } }}
                    >
                      <Checkbox
                        checked={selected.has(o.serverRelativeUrl)}
                        onChange={(): void => this._toggleFile(o.serverRelativeUrl)}
                        disabled={recycling}
                        styles={{ root: { paddingTop: 4 } }}
                      />
                      <Stack grow tokens={{ childrenGap: 2 }}>
                        <Text>
                          <strong>{o.fileName}</strong>{' '}
                          <Text variant="small" styles={{ root: { color: 'var(--neutralSecondary, #605e5c)' } }}>
                            in folder <code>{o.pageFolder}</code> • {fmtBytes(o.size)} • modified{' '}
                            {o.modified ? new Date(o.modified).toLocaleDateString() : 'unknown'}
                          </Text>
                        </Text>
                        <a
                          href={o.serverRelativeUrl}
                          {...EXTERNAL_LINK_ATTRS}
                          style={{ ...LINK_STYLE, ...MONO_STYLE }}
                        >
                          {o.serverRelativeUrl}
                        </a>
                      </Stack>
                    </Stack>
                  ))}
                </Stack>
              )}
            </Stack>
          );
        })}

        <Text variant="small" styles={{ root: { color: 'var(--neutralSecondary, #605e5c)', marginTop: 16 } }}>
          Recycle moves files to the SharePoint site recycle bin (~93-day retention: 30 days
          first-stage + 63 days second-stage). Recovery: open the site, Settings → Site contents →
          Recycle bin. Action runs as you, so the recycle-bin record shows your name.
        </Text>
      </Stack>
    );
  }
}
