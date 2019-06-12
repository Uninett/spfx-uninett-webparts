import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'UserDirectoryWebPartStrings';
import UserDirectory from './components/UserDirectory';
import { IUserDirectoryProps } from './components/IUserDirectoryProps';

export interface IUserDirectoryWebPartProps {
  api: string;
}

export default class UserDirectoryWebPart extends BaseClientSideWebPart<IUserDirectoryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IUserDirectoryProps > = React.createElement(
      UserDirectory,
      {
        context: this.context,
        api: this.properties.api
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
