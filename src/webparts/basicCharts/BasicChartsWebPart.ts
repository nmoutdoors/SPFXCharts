import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'BasicChartsWebPartStrings';
import BasicCharts from './components/BasicCharts';
import { IBasicChartsProps } from './components/IBasicChartsProps';

export interface IBasicChartsWebPartProps {
  description: string;
}

export default class BasicChartsWebPart extends BaseClientSideWebPart<IBasicChartsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBasicChartsProps> = React.createElement(
      BasicCharts,
      {
        description: this.properties.description,
        isDarkTheme: false,
        environmentMessage: '',
        hasTeamsContext: false,
        userDisplayName: '',
        context: this.context
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


