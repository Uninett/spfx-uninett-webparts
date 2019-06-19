import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'SearchBoxWebPartStrings';
import SearchBox from './components/SearchBox';
import { ISearchBoxProps } from './components/ISearchBoxProps';

export interface ISearchBoxWebPartProps {
  searchBoxLabel: string;
}

export default class SearchBoxWebPart extends BaseClientSideWebPart<ISearchBoxWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISearchBoxProps > = React.createElement(
      SearchBox,
      {
        searchBoxLabel: this.properties.searchBoxLabel
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
                PropertyPaneTextField('searchBoxLabel', {
                  label: strings.SearchBoxPropLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
