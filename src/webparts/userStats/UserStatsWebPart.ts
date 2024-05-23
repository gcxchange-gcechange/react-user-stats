import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'UserStatsWebPartStrings';
import UserStats from './components/UserStats';
import { IUserStatsProps } from './components/IUserStatsProps';

export interface IUserStatsWebPartProps {
  description: string;
  storageCapacity: number;
  storageUnit: string;
}

export default class UserStatsWebPart extends BaseClientSideWebPart<IUserStatsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IUserStatsProps> = React.createElement(
      UserStats,
      {
        description: this.properties.description,
        context: this.context,
        storageCapacity: this.properties.storageCapacity,
        storageUnit: this.properties.storageUnit
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

  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                }),
                PropertyPaneTextField('storageCapacity', {
                  label: strings.StorageCapacityLabel
                }),
                PropertyPaneChoiceGroup('storageUnit', {
                  label: strings.StorageUnitLabel,
                  options: [
                    { text: "GB", key: "GB" },
                    { text: "TB", key: "TB" }
                ]                  
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
