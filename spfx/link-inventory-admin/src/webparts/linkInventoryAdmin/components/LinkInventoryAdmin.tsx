import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import type { SPHttpClient, AadHttpClientFactory, HttpClient } from '@microsoft/sp-http';
import { LinkInventoryService } from '../services/LinkInventoryService';
import { LinkInventoryTab } from './LinkInventoryTab';
import { DuplicatesTab } from './DuplicatesTab';
import { HelpTab } from './HelpTab';

export interface ILinkInventoryAdminProps {
  spHttpClient: SPHttpClient;
  httpClient: HttpClient;
  aadHttpClientFactory: AadHttpClientFactory;
  siteUrl: string;
  /** Tenant origin derived from siteUrl, e.g. "https://contoso.sharepoint.com" — used to resolve server-relative paths in DuplicatesTab. */
  tenantOrigin: string;
  functionUrl: string;
  /** Entra app api:// URI for OBO. */
  entraAppApiUri: string;
}

interface ILinkInventoryAdminState {
  error: string;
  selectedTab: string;
}

export class LinkInventoryAdmin extends React.Component<ILinkInventoryAdminProps, ILinkInventoryAdminState> {
  private linkInventoryService: LinkInventoryService;

  public constructor(props: ILinkInventoryAdminProps) {
    super(props);
    this.state = { error: '', selectedTab: 'inventory' };

    this.linkInventoryService = new LinkInventoryService(
      props.aadHttpClientFactory,
      props.functionUrl,
      props.entraAppApiUri,
    );
  }

  private _onError = (message: string): void => {
    this.setState({ error: message });
  };

  private _dismissError = (): void => {
    this.setState({ error: '' });
  };

  public render(): React.ReactElement<ILinkInventoryAdminProps> {
    const configMissing = !this.props.functionUrl || !this.props.entraAppApiUri;

    return (
      <div>
        {configMissing && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Web part is not configured. Open the property pane and set both the
            Azure Function URL and the Entra app api:// URI.
          </MessageBar>
        )}
        {this.state.error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={this._dismissError}
            dismissButtonAriaLabel="Close"
          >
            {this.state.error}
          </MessageBar>
        )}
        <Pivot
          selectedKey={this.state.selectedTab}
          onLinkClick={(item?: PivotItem) => {
            if (item?.props.itemKey) this.setState({ selectedTab: item.props.itemKey });
          }}
        >
          <PivotItem headerText="Link Inventory" itemKey="inventory" itemIcon="Search">
            <LinkInventoryTab
              service={this.linkInventoryService}
              onError={this._onError}
            />
          </PivotItem>
          <PivotItem headerText="Duplicates" itemKey="duplicates" itemIcon="DocumentSet">
            <DuplicatesTab
              service={this.linkInventoryService}
              tenantOrigin={this.props.tenantOrigin}
              onError={this._onError}
            />
          </PivotItem>
          <PivotItem headerText="Help" itemKey="help" itemIcon="Help">
            <HelpTab />
          </PivotItem>
        </Pivot>
      </div>
    );
  }
}
