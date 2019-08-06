import { SPHttpClient } from '@microsoft/sp-http';
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

import * as strings from 'OrderGroupWebPartStrings';
import OrderGroup from './components/OrderGroup';
import { IOrderGroupProps } from './components/IOrderGroupProps';
import { initializeIcons } from '@uifabric/icons';
import { SiteType } from './interfaces/SiteType';
initializeIcons();

export interface IOrderGroupWebPartProps {
  description: string;
  departmentSiteTypeName: SiteType;
  hideParentDepartment: boolean;
}

export default class OrderGroupWebPart extends BaseClientSideWebPart<IOrderGroupWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IOrderGroupProps > = React.createElement(
      OrderGroup,
      {
        description: this.properties.description,
        context: this.context,
        departmentSiteTypeName: this.properties.departmentSiteTypeName,
        hideParentDepartment: this.properties.hideParentDepartment
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneChoiceGroup('departmentSiteTypeName', {
                  label: strings.PropertyPaneDepartmentSiteTypeName,
                  options: [
                   { key: SiteType.department, text: 'Department' },
                   { key: SiteType.section, text: 'Section', checked: true }
                 ]
               }),
               PropertyPaneCheckbox('hideParentDepartment', {
                  text: strings.PropertyPaneHideParentDepartment,
                  checked: false  
              })
              ]
            }
          ]
        }
      ]
    };
  }
}
