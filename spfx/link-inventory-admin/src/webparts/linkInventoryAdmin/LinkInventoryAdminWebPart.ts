import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { LinkInventoryAdmin, type ILinkInventoryAdminProps } from './components/LinkInventoryAdmin';
import * as strings from 'LinkInventoryAdminWebPartStrings';

export interface ILinkInventoryAdminWebPartProps {
  /** Base URL of the deployed Azure Function, e.g. https://func-link-inventory.azurewebsites.net */
  functionUrl: string;
  /** Entra app api:// URI used as the OBO target resource, e.g. api://12345678-1234-1234-1234-123456789012 */
  entraAppApiUri: string;
}

export default class LinkInventoryAdminWebPart extends BaseClientSideWebPart<ILinkInventoryAdminWebPartProps> {

  public render(): void {
    const siteAbsolute = this.context.pageContext.web.absoluteUrl;
    const tenantOrigin = (() => {
      try { return new URL(siteAbsolute).origin; } catch { return ''; }
    })();

    const element: React.ReactElement<ILinkInventoryAdminProps> = React.createElement(
      LinkInventoryAdmin,
      {
        spHttpClient: this.context.spHttpClient,
        httpClient: this.context.httpClient,
        aadHttpClientFactory: this.context.aadHttpClientFactory,
        siteUrl: siteAbsolute,
        tenantOrigin,
        functionUrl: this.properties.functionUrl || '',
        entraAppApiUri: this.properties.entraAppApiUri || '',
      },
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('functionUrl', {
                  label: strings.FunctionUrlLabel,
                  description: 'Base URL of your deployed Azure Function (no trailing slash)',
                }),
                PropertyPaneTextField('entraAppApiUri', {
                  label: strings.EntraAppApiUriLabel,
                  description: 'api:// URI of the Entra app the function uses for OBO. Find it on the app registration\'s "Expose an API" page.',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
