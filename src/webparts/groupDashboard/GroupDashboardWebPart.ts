import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GroupDashboardWebPartStrings';
import GroupDashboard from './components/GroupDashboard';
import { IGroupDashboardProps } from './components/IGroupDashboardProps';
import { sp } from '@pnp/sp';
import PropertyPaneAccessControl from "property-pane-access-control";

export interface IGroupDashboardWebPartProps {
  description: string;
}

export default class GroupDashboardWebPart extends BaseClientSideWebPart<IGroupDashboardWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGroupDashboardProps> = React.createElement(
      GroupDashboard,
      {
        description: this.properties.description,
        context: this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await sp.setup({
      spfxContext: this.context,
      defaultCachingStore: "session",
      defaultCachingTimeoutSeconds: 60,
      globalCacheDisable: false,
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupFields: [
                PropertyPaneAccessControl("description", {
                  key: "accesscontrol",
                  permissions: [
                    "view",
                    "edit",
                    "create",
                    "delete"
                  ],
                  context: this.context,
                }),
              ]
            },
          ]
        }
      ]
    };
  }
}
