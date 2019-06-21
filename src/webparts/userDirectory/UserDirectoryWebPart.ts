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
  PropertyPaneButtonType,
  PropertyPaneChoiceGroup,
  PropertyPaneLabel
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
  useBuiltInSearch: boolean;
  searchBoxPlaceholder: string;
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
    
    // Sets column headers
    this.setColumnHeaders();
    
    if (!this.properties.hasBeenInitialised) this.setDefaultPlaceholder();

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
        useBuiltInSearch: this.properties.useBuiltInSearch,
        searchBoxPlaceholder: this.properties.searchBoxPlaceholder,
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

  private resetColumnHeaders() {
    this.clearCustomHeaders();
    this.setColumnHeaders();
    this.context.propertyPane.refresh();
    this.render();
  }

  private clearCustomHeaders() {
    this.properties.customName = "";
    this.properties.customJobTitle = "";
    this.properties.customDepartment = "";
    this.properties.customOfficeLocation = "";
    this.properties.customCity = "";
    this.properties.customPhone = "";
    this.properties.customMail = "";
  }

  private setColumnHeaders() {
    // Sets custom column headers, or default headers if custom is empty
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

  private setDefaultPlaceholder() {
    this.properties.searchBoxPlaceholder = strings.DefaultSearchPlaceholder;
    this.properties.hasBeenInitialised = true;
    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    let searchBoxProps: any;

    if (this.properties.useBuiltInSearch) {
      searchBoxProps = PropertyPaneTextField('searchBoxPlaceholder', {
        label: strings.SearchBoxPlaceholderLabel
      });
    }
    else {
      searchBoxProps = PropertyPaneLabel('lblSearchHelp',{
        text: strings.SearchBoxHelpLabel
      });
    }

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
                  text: strings.ApplyApiButton,
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
            {
              groupName: strings.GroupSearchBox,
              groupFields: [
                PropertyPaneToggle('useBuiltInSearch', {
                  label: strings.UseBuiltInSearchLabel,
                  checked: true,
                  onText: strings.UseBuiltInSearchOn,
                  offText: strings.UseBuiltInSearchOff
                }),
                // Property dependent on toggle choice
                searchBoxProps
              ]
            }                   
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
                  placeholder: strings.CustomHeaderPlaceholder,
                  disabled: !this.properties.showName
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showJobTitle', {                
                  text: strings.ColJobTitle,
                  checked: true
                }),
                PropertyPaneTextField('customJobTitle',{
                  placeholder: strings.CustomHeaderPlaceholder,
                  disabled: !this.properties.showJobTitle
                }),
                PropertyPaneHorizontalRule(),     
                PropertyPaneCheckbox('showDepartment', {                
                  text: strings.ColDepartment,
                  checked: true
                }),
                PropertyPaneTextField('customDepartment',{
                  placeholder: strings.CustomHeaderPlaceholder,
                  disabled: !this.properties.showDepartment
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showOfficeLocation', {                
                  text: strings.ColOfficeLocation,
                  checked: false
                }),
                PropertyPaneTextField('customOfficeLocation',{
                  placeholder: strings.CustomHeaderPlaceholder,
                  disabled: !this.properties.showOfficeLocation
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showCity', {                
                  text: strings.ColCity,
                  checked: false
                }),      
                PropertyPaneTextField('customCity',{
                  placeholder: strings.CustomHeaderPlaceholder,
                  disabled: !this.properties.showCity
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showPhone', {                
                  text: strings.ColPhone,
                  checked: true
                }),
                PropertyPaneTextField('customPhone',{
                  placeholder: strings.CustomHeaderPlaceholder,
                  disabled: !this.properties.showPhone
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneCheckbox('showMail', {                
                  text: strings.ColMail,
                  checked: true
                }),
                PropertyPaneTextField('customMail',{
                  placeholder: strings.CustomHeaderPlaceholder,
                  disabled: !this.properties.showMail
                }),
                PropertyPaneButton('btnReset', {
                  text: strings.BtnResetText,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.resetColumnHeaders.bind(this)
                 })
              ]
            }
          ]            
        }
      ]
    };
  }
}
