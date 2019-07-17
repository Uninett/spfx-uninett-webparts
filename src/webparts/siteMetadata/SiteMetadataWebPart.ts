import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-webpart-base';
import * as strings from 'SiteMetadataWebPartStrings';
import Sitemetadata from './components/SiteMetadata';
import { ISiteMetadataProps } from './components/ISiteMetadataProps';

export interface ISiteMetadataWebPartProps {
  description: string;
  orderSiteURL: string;
  editName: boolean;
  editDescription: boolean;
  editProjectGoal: boolean;
  editProjectPurpose: boolean;
  editOwner: boolean;
  editShortName: boolean;
  editParentDepartment: boolean;
  // editOwningDepartment: boolean;
  editProjectNumber: boolean;
  editStartDate: boolean;
  editEndDate: boolean;
}

export default class SiteMetadataWebPart extends BaseClientSideWebPart<ISiteMetadataWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISiteMetadataProps > = React.createElement(
      Sitemetadata,
      {
        description: this.properties.description,
        orderSiteURL: this.properties.orderSiteURL,
        editName: this.properties.editName,
        editDescription: this.properties.editDescription,
        editProjectGoal: this.properties.editProjectGoal,
        editProjectPurpose: this.properties.editProjectPurpose,
        editOwner: this.properties.editOwner,
        editShortName: this.properties.editShortName,
        editParentDepartment: this.properties.editParentDepartment,
        // editOwningDepartment: this.properties.editOwningDepartment,
        editProjectNumber: this.properties.editProjectNumber,
        editStartDate: this.properties.editStartDate,
        editEndDate: this.properties.editEndDate,
        context: this.context,
        displayMode: this.displayMode
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
                PropertyPaneTextField('orderSiteURL', {
                  label: strings.OrderSiteURLLabel
                }),
                PropertyPaneCheckbox('editName', {text: strings.editName}),
                PropertyPaneCheckbox('editDescription', {text: strings.editDescription}),
                PropertyPaneCheckbox('editProjectGoal', {text: strings.editProjectGoal}),
                PropertyPaneCheckbox('editProjectPurpose', {text: strings.editProjectPurpose}),
                PropertyPaneCheckbox('editOwner', {text: strings.editOwner}),
                PropertyPaneCheckbox('editShortName', {text: strings.editShortName}),
                PropertyPaneCheckbox('editParentDepartment', {text: strings.editParentDepartment}),
                // PropertyPaneCheckbox('editOwningDepartment', {text: strings.editOwningDepartment}),
                PropertyPaneCheckbox('editProjectNumber', {text: strings.editProjectNumber}),
                PropertyPaneCheckbox('editStartDate', {text: strings.editStartDate}),
                PropertyPaneCheckbox('editEndDate', {text: strings.editEndDate}),
              ]
            }
          ]
        }
      ]
    };
  }
}
