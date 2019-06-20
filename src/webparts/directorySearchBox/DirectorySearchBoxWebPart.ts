import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'DirectorySearchBoxWebPartStrings';
import DirectorySearchBox from './components/DirectorySearchBox';
import { IDirectorySearchBoxProps } from './components/IDirectorySearchBoxProps';

export interface IDirectorySearchBoxWebPartProps {
  searchBoxPlaceholder: string;
  hasBeenInitialised: boolean;
}

export default class DirectorySearchBoxWebPart extends BaseClientSideWebPart<IDirectorySearchBoxWebPartProps> {

  public render(): void {

    if (!this.properties.hasBeenInitialised) this.setDefaultPlaceholder();

    const element: React.ReactElement<IDirectorySearchBoxProps > = React.createElement(
      DirectorySearchBox,
      {
        searchBoxPlaceholder: this.properties.searchBoxPlaceholder
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

  private setDefaultPlaceholder() {
    this.properties.searchBoxPlaceholder = strings.DefaultPlaceholder;
    this.properties.hasBeenInitialised = true;
    this.context.propertyPane.refresh();
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
                PropertyPaneTextField('searchBoxPlaceholder', {
                  label: strings.PlaceholderLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
