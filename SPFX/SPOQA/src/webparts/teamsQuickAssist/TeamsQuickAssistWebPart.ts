import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TeamsQuickAssistWebPartStrings';
import TeamsQuickAssist from './components/TeamsQuickAssist';
import { ITeamsQuickAssistProps } from './components/ITeamsQuickAssistProps';

export interface ITeamsQuickAssistWebPartProps {
  description: string;
}

export default class TeamsQuickAssistWebPart extends BaseClientSideWebPart<ITeamsQuickAssistWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITeamsQuickAssistProps> = React.createElement(
      TeamsQuickAssist,
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
