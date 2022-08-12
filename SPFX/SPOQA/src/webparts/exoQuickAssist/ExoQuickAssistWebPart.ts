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

export interface IExoQuickAssistWebPartProps {
  description: string;
}

export default class ExoQuickAssistWebPart extends BaseClientSideWebPart<IExoQuickAssistWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IExoQuickAssistProps> = React.createElement(
      ExoQuickAssist,
      {
        description: this.properties.description
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
