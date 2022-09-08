import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ExoQuickAssistWebPartStrings';
import ExoQuickAssist from './components/ExoQuickAssist';
import { IExoQuickAssistProps } from './components/IExoQuickAssistProps';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as contextHelper from '../Helpers/ContextHelper';
import {MSGraphClient, SPHttpClient} from '@microsoft/sp-http';

export interface IExoQuickAssistWebPartProps {
  description: string;
}

export default class ExoQuickAssistWebPart extends BaseClientSideWebPart<IExoQuickAssistWebPartProps> {
  
  private graphClient: MSGraphClient;

  constructor() {
    super();
    contextHelper.default.SetInstace(this.context);
    SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/6.0.0/css/fabric-6.0.0.scoped.css');
  }

  public render(): void {
    const element: React.ReactElement<IExoQuickAssistProps> = React.createElement(
      ExoQuickAssist,
      {
         msGraphClient: this.graphClient,        
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
