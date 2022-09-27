import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import TeamsQuickAssist from './components/TeamsQuickAssist';
import { ITeamsQuickAssistProps } from './components/ITeamsQuickAssistProps';

import { SPComponentLoader } from '@microsoft/sp-loader';
import * as contextHelper from '../Helpers/ContextHelper';
import {MSGraphClient, SPHttpClient} from '@microsoft/sp-http';

export interface ITeamsQuickAssistWebPartProps {
  description: string;
}

export default class TeamsQuickAssistWebPart extends BaseClientSideWebPart<ITeamsQuickAssistWebPartProps> {
  private graphClient: MSGraphClient;

  constructor() {
    super();
    contextHelper.default.SetInstace(this.context);
    SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/6.0.0/css/fabric-6.0.0.scoped.css');
  }

  public onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          this.graphClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  public render(): void {
    const element: React.ReactElement<ITeamsQuickAssistProps> = React.createElement(
      TeamsQuickAssist,
      {
        msGraphClient: this.graphClient,        
        currentUser:this.context.pageContext.user,
        ctx:this.context
      }
    );

    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        // get information about the current user from the Microsoft Graph
        client
          .api('/me')
          .get((error, response: any, rawResponse?: any) => {
            console.log(`${response} and ${error}`);
        });
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
