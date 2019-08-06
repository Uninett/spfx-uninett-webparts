import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import * as strings from 'SiteCatalogWebPartStrings';
import SiteCatalog from './components/siteCatalog/SiteCatalog';
import { ISiteCatalogProps } from './components/siteCatalog/ISiteCatalogProps';

export interface ISiteCatalogWebPartProps {
  listRows: string;
  showNewButton: boolean;
  siteTypes: string;
  hideParentDepartment: boolean;
  searchSetting: SearchType;
}

export default class SiteCatalogWebPart extends BaseClientSideWebPart<ISiteCatalogWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISiteCatalogProps > = React.createElement(
      SiteCatalog,
      {
        listRows: this.properties.listRows,
        showNewButton: this.properties.showNewButton,
        context: this.context,
        siteTypes: this.properties.siteTypes,
        hideParentDepartment: this.properties.hideParentDepartment,
        searchType: this.properties.searchSetting
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('listRows', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneCheckbox('showNewButton', {
                  text: strings.ShowNewButtonLabel,
                  checked: true
                }),
                PropertyPaneCheckbox('hideParentDepartment', {
                  text: strings.PropertyPaneHideParentDepartment,
                  checked: false
                }),
                PropertyPaneTextField('siteTypes', {
                  label: strings.PropertyPaneSiteTypes
                })
              ]
            },
            {
              groupName: strings.PropertyPaneSearchSetting,
              groupFields: [
                PropertyPaneChoiceGroup('searchSetting', {
                  label: strings.PropertyPaneSearchType,
                  options: [
                    {
                      key: 'graph',
                      text: strings.PropertyPaneSearchTypeGraph
                    },
                    {
                      key: 'javascript',
                      text: strings.PropertyPaneSearchTypeJavascript,
                      checked: true
                    }
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
