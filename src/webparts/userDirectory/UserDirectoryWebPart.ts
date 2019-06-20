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
  hasBeenInitialised: boolean;
  api: string;
  isApiChanged: boolean;
  compactMode: boolean;
  alternatingColours: boolean;
  showPhoto: boolean;
  showName: boolean;
  showJobTitle: boolean;
  showDepartment: boolean;
  showOfficeLocation: boolean;
  showCity: boolean;
  showPhone: boolean;
  showMail: boolean;
  colName: string;
  colJobTitle: string;
  colDepartment: string;
  colOfficeLocation: string;
  colCity: string;
  colPhone: string;
  colMail: string;
  customName: string;
  customJobTitle: string;
  customDepartment: string;
  customOfficeLocation: string;
  customCity: string;
  customPhone: string;
  customMail: string;
}

export default class UserDirectoryWebPart extends BaseClientSideWebPart<IUserDirectoryWebPartProps> {
  

  public render(): void {
    
    // Sets column titles
    this.setColumnTitles();
    
    // Reloads entire compononent if data source API is updated
    if (this.properties.isApiChanged) {
      ReactDom.unmountComponentAtNode(this.domElement);
      this.properties.isApiChanged = false;
    }

    const element: React.ReactElement<IUserDirectoryProps > = React.createElement(
      UserDirectory,
      {
        context: this.context,
        api: this.properties.api,
        isApiChanged: this.properties.isApiChanged,
        compactMode: this.properties.compactMode,
        alternatingColours: this.properties.alternatingColours,
        showPhoto: this.properties.showPhoto,
        showName: this.properties.showName,
        showJobTitle: this.properties.showJobTitle,
        showDepartment: this.properties.showDepartment,
        showOfficeLocation: this.properties.showOfficeLocation,
        showCity: this.properties.showCity,
        showPhone: this.properties.showPhone,
        showMail: this.properties.showMail,
        colName: this.properties.colName,
        colJobTitle: this.properties.colJobTitle,
        colDepartment: this.properties.colDepartment,
        colOfficeLocation: this.properties.colOfficeLocation,
        colCity: this.properties.colCity,
        colPhone: this.properties.colPhone,
        colMail: this.properties.colMail
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
    this.clearCustomTitles();
    this.setColumnTitles();
    //this.setDefaultColumnTitles();
    this.context.propertyPane.refresh();
    this.render();
  }

  private setDefaultColumnTitles() {
    this.properties.colName = strings.ColName;
    this.properties.colJobTitle = strings.ColJobTitle;
    this.properties.colDepartment = strings.ColDepartment;
    this.properties.colOfficeLocation = strings.ColOfficeLocation;
    this.properties.colCity = strings.ColCity;
    this.properties.colPhone = strings.ColPhone;
    this.properties.colMail = strings.ColMail;
  }

  private clearCustomTitles() {
    this.properties.customName = "";
    this.properties.customJobTitle = "";
    this.properties.customDepartment = "";
    this.properties.customOfficeLocation = "";
    this.properties.customCity = "";
    this.properties.customPhone = "";
    this.properties.customMail = "";
  }

  private setColumnTitles() {
    // Sets custom column titles, or default titles if custom is empty
    this.properties.colName = this.properties.customName || strings.ColName;
    this.properties.colJobTitle = this.properties.customJobTitle || strings.ColJobTitle;
    this.properties.colDepartment = this.properties.customDepartment || strings.ColDepartment;
    this.properties.colOfficeLocation = this.properties.customOfficeLocation || strings.ColOfficeLocation;
    this.properties.colCity = this.properties.customCity || strings.ColCity;
    this.properties.colPhone = this.properties.customPhone || strings.ColPhone;
    this.properties.colMail = this.properties.customMail || strings.ColMail;
  }

  private updateApi() {
    this.properties.isApiChanged = true;
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
                  label: strings.ApiLabel
                }),
                PropertyPaneButton('btnApplyApi', {
                  text: "Apply",
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.updateApi.bind(this)
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
                  text: strings.ColPhoto,
                  checked: true
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showName', {                
                  text: strings.ColName,
                  checked: true
                }),
                PropertyPaneTextField('customName',{
                  placeholder: strings.CustomTitlePlaceholder,
                  disabled: !this.properties.showName
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showJobTitle', {                
                  text: strings.ColJobTitle,
                  checked: true
                }),
                PropertyPaneTextField('customJobTitle',{
                  placeholder: strings.CustomTitlePlaceholder,
                  disabled: !this.properties.showJobTitle
                }),
                PropertyPaneHorizontalRule(),     
                PropertyPaneCheckbox('showDepartment', {                
                  text: strings.ColDepartment,
                  checked: true
                }),
                PropertyPaneTextField('customDepartment',{
                  placeholder: strings.CustomTitlePlaceholder,
                  disabled: !this.properties.showDepartment
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showOfficeLocation', {                
                  text: strings.ColOfficeLocation,
                  checked: false
                }),
                PropertyPaneTextField('customOfficeLocation',{
                  placeholder: strings.CustomTitlePlaceholder,
                  disabled: !this.properties.showOfficeLocation
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showCity', {                
                  text: strings.ColCity,
                  checked: false
                }),      
                PropertyPaneTextField('customCity',{
                  placeholder: strings.CustomTitlePlaceholder,
                  disabled: !this.properties.showCity
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showPhone', {                
                  text: strings.ColPhone,
                  checked: true
                }),
                PropertyPaneTextField('customPhone',{
                  placeholder: strings.CustomTitlePlaceholder,
                  disabled: !this.properties.showPhone
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showMail', {                
                  text: strings.ColMail,
                  checked: true
                }),
                PropertyPaneTextField('customMail',{
                  placeholder: strings.CustomTitlePlaceholder,
                  disabled: !this.properties.showMail
                }),
                PropertyPaneButton('btnReset', {
                  text: strings.BtnResetText,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.resetColumnTitles.bind(this)
                 })
              ]
            }/*,
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
            }*/
          ]
            
        }
      ]
    };
  }
}
