import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
import SharePointOnlineQuickAssist from './components/SharePointOnlineQuickAssist';
import { ISharePointOnlineQuickAssistProps } from './components/ISharePointOnlineQuickAssistProps';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface ISharePointOnlineQuickAssistWebPartProps {
  description: string;
}

export default class SharePointOnlineQuickAssistWebPart extends BaseClientSideWebPart<ISharePointOnlineQuickAssistWebPartProps> {
  constructor() {
    super();
    SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/6.0.0/css/fabric-6.0.0.scoped.css');
  }
  
  public render(): void {
    const element: React.ReactElement<ISharePointOnlineQuickAssistProps> = React.createElement(
      SharePointOnlineQuickAssist,
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
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
