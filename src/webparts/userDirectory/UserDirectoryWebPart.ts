import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';

import * as strings from 'UserDirectoryWebPartStrings';
import UserDirectory from './components/UserDirectory';
import { IUserDirectoryProps } from './components/IUserDirectoryProps';

export interface IUserDirectoryWebPartProps {
  api: string;
  showPhoto: boolean;
  showJobTitle: boolean;
  showMail: boolean;
  showPhone: boolean;
  compactMode: boolean;
}

export default class UserDirectoryWebPart extends BaseClientSideWebPart<IUserDirectoryWebPartProps> {

  public render(): void {
    console.log("compact mode on render: " + this.properties.compactMode);
    const element: React.ReactElement<IUserDirectoryProps > = React.createElement(
      UserDirectory,
      {
        context: this.context,
        api: this.properties.api,
        showPhoto: this.properties.showPhoto,
        showJobTitle: this.properties.showJobTitle,
        showMail: this.properties.showMail,
        showPhone: this.properties.showPhone,
        compactMode: this.properties.compactMode
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
                PropertyPaneTextField('api', {
                  label: strings.ApiLabel,
                  value: "users"
                }),                
                PropertyPaneToggle('compactMode', {
                  label: strings.CompactModeLabel,
                  checked: false,
                  onText:"Compact",
                  offText:"Normal"
                })
              ]
            },
            {
              groupName: "Select columns to display",
              groupFields: [
                PropertyPaneCheckbox('showPhoto', {                
                  text: "Photo",
                  checked: true
                }),
                PropertyPaneCheckbox('showJobTitle', {                
                  text: "Job Title",
                  checked: true
                }),
                PropertyPaneCheckbox('showMail', {                
                  text: "Mail",
                  checked: true
                }),
                PropertyPaneCheckbox('showPhone', {                
                  text: "Phone",
                  checked: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
