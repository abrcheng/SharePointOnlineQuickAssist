import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
import SharePointOnlineQuickAssist from './components/SharePointOnlineQuickAssist';
import { ISharePointOnlineQuickAssistProps } from './components/ISharePointOnlineQuickAssistProps';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as contextHelper from '../Helpers/ContextHelper';
import {MSGraphClient, SPHttpClient} from '@microsoft/sp-http';

export interface ISharePointOnlineQuickAssistWebPartProps {
  description: string;
 }

export default class SharePointOnlineQuickAssistWebPart extends BaseClientSideWebPart<ISharePointOnlineQuickAssistWebPartProps> {
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
    const element: React.ReactElement<ISharePointOnlineQuickAssistProps> = React.createElement(
      SharePointOnlineQuickAssist,
      {        
        msGraphClient: this.graphClient,
        spHttpClient:this.context.spHttpClient,
        webAbsoluteUrl:this.context.pageContext.web.absoluteUrl,
        webUrl:this.context.pageContext.legacyPageContext["webServerRelativeUrl"],
        rootUrl:this.context.pageContext.site.absoluteUrl.indexOf(".sharepoint.com") >1 ? this.context.pageContext.site.absoluteUrl.substring(0,this.context.pageContext.site.absoluteUrl.indexOf(".sharepoint.com")+(".sharepoint.com").length): this.context.pageContext.site.absoluteUrl.substring(0,this.context.pageContext.site.absoluteUrl.indexOf(".sharepoint.cn")+(".sharepoint.cn").length),
        currentUser:this.context.pageContext.user,
        ctx:this.context
      }     
    );

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
