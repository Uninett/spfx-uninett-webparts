import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneCheckbox,
  PropertyPaneHorizontalRule,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';

import * as strings from 'UserDirectoryWebPartStrings';
import UserDirectory from './components/UserDirectory';
import { IUserDirectoryProps } from './components/IUserDirectoryProps';

export interface IUserDirectoryWebPartProps {
  api: string;
  compactMode: boolean;
  alternatingColours: boolean;
  showPhoto: boolean;
  showJobTitle: boolean;
  showDepartment: boolean;
  showOfficeLocation: boolean;
  showCity: boolean;
  showPhone: boolean;
  showMail: boolean;
  colNameTitle: string;
  colJobTitleTitle: string;
  colDepartmentTitle: string;
  colOfficeLocationTitle: string;
  colCityTitle: string;
  colPhoneTitle: string;
  colMailTitle: string;
}

export default class UserDirectoryWebPart extends BaseClientSideWebPart<IUserDirectoryWebPartProps> {
  private _hasBeenInitialised: boolean = false;


  public render(): void {

    if (!this._hasBeenInitialised) {
      this.resetPropertyStrings();
      this._hasBeenInitialised = true;
    }

    const element: React.ReactElement<IUserDirectoryProps > = React.createElement(
      UserDirectory,
      {
        context: this.context,
        api: this.properties.api,
        compactMode: this.properties.compactMode,
        alternatingColours: this.properties.alternatingColours,
        showPhoto: this.properties.showPhoto,
        showJobTitle: this.properties.showJobTitle,
        showDepartment: this.properties.showDepartment,
        showOfficeLocation: this.properties.showOfficeLocation,
        showCity: this.properties.showCity,
        showPhone: this.properties.showPhone,
        showMail: this.properties.showMail,
        colNameTitle: this.properties.colNameTitle,
        colJobTitleTitle: this.properties.colJobTitleTitle,
        colDepartmentTitle: this.properties.colDepartmentTitle,
        colOfficeLocationTitle: this.properties.colOfficeLocationTitle,
        colCityTitle: this.properties.colCityTitle,
        colPhoneTitle: this.properties.colPhoneTitle,
        colMailTitle: this.properties.colMailTitle
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

  private resetColumnTitles() {
    this.resetPropertyStrings();
    this.context.propertyPane.refresh();
    this.render();
  }

  private resetPropertyStrings() {
    this.properties.colNameTitle = strings.ColNameTitle;
    this.properties.colJobTitleTitle = strings.ColJobTitleTitle;
    this.properties.colDepartmentTitle = strings.ColDepartmentTitle;
    this.properties.colOfficeLocationTitle = strings.ColOfficeLocationTitle;
    this.properties.colCityTitle = strings.ColCityTitle;
    this.properties.colPhoneTitle = strings.ColPhoneTitle;
    this.properties.colMailTitle = strings.ColMailTitle;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    return {
      pages: [
        {
          header: {
            description: strings.PageGeneralDescription
          },
          groups: [
            {
              groupName: strings.GroupDataSource,
              groupFields: [
                PropertyPaneTextField('api', {
                  label: strings.ApiLabel,
                  value: "users"
                })
              ]
            },
            {
              groupName: strings.GroupAppearance,
              groupFields: [
                PropertyPaneToggle('compactMode', {
                  label: strings.CompactModeLabel,
                  checked: false,
                  onText: strings.CompactModeOn,
                  offText: strings.CompactModeOff
                }),
                PropertyPaneToggle('alternatingColours', {
                  label: strings.AlternateColoursLabel,
                  checked: false,
                  onText: strings.AlternateColoursOn,
                  offText: strings.AlternateColoursOff
                })
              ]
            },
            
          ]
        },
        {
          header: {
            description: strings.PageColumnsDescription
          },
          groups: [
            {
              groupName: strings.GroupColumns,
              groupFields: [
                PropertyPaneCheckbox('showPhoto', {                
                  text: strings.ColPhotoTitle,
                  checked: true
                }),
                PropertyPaneCheckbox('showJobTitle', {                
                  text: strings.ColJobTitleTitle,
                  checked: true
                }),         
                PropertyPaneCheckbox('showDepartment', {                
                  text: strings.ColDepartmentTitle,
                  checked: true
                }),
                PropertyPaneCheckbox('showOfficeLocation', {                
                  text: strings.ColOfficeLocationTitle,
                  checked: false
                }),
                PropertyPaneCheckbox('showCity', {                
                  text: strings.ColCityTitle,
                  checked: false
                }),      
                PropertyPaneCheckbox('showPhone', {                
                  text: strings.ColPhoneTitle,
                  checked: true
                }),
                PropertyPaneCheckbox('showMail', {                
                  text: strings.ColMailTitle,
                  checked: true
                })
              ]
            },
            {
              groupName: strings.GroupColumnTitles,
              groupFields: [
                PropertyPaneTextField('colNameTitle',{
                  placeholder: strings.ColNameTitle,
                }),
                PropertyPaneTextField('colJobTitleTitle',{
                  placeholder: strings.ColJobTitleTitle,
                  disabled: !this.properties.showJobTitle
                }),
                PropertyPaneTextField('colDepartmentTitle',{
                  placeholder: strings.ColDepartmentTitle,
                  disabled: !this.properties.showDepartment
                }),
                PropertyPaneTextField('colOfficeLocationTitle',{
                  placeholder: strings.ColOfficeLocationTitle,
                  disabled: !this.properties.showOfficeLocation
                }),
                PropertyPaneTextField('colCityTitle',{
                  placeholder: strings.ColCityTitle,
                  disabled: !this.properties.showCity
                }),
                PropertyPaneTextField('colPhoneTitle',{
                  placeholder: strings.ColPhoneTitle,
                  disabled: !this.properties.showPhone
                }),
                PropertyPaneTextField('colMailTitle',{
                  placeholder: strings.ColMailTitle,
                  disabled: !this.properties.showMail
                }),
                PropertyPaneButton('btnReset', {
                  text: strings.BtnResetText,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.resetColumnTitles.bind(this)
                 })
              ]
            }
          ]
            
        }
      ]
    };
  }
}
