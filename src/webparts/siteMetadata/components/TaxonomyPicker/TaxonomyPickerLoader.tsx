import * as React from 'react';
import { ITaxonomyPickerLoaderProps } from './ITaxonomyPickerLoaderProps';
import { ITaxonomyPickerLoaderState } from './ITaxonomyPickerLoaderState';
import { SPComponentLoader } from '@microsoft/sp-loader';

import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

import TaxonomyPicker, { ITaxonomyPickerProps } from "react-taxonomypicker";
import "react-taxonomypicker/dist/React.TaxonomyPicker.css";
import * as strings from 'SiteMetadataWebPartStrings';

export default class TaxonomyPickerLoader extends React.Component<ITaxonomyPickerLoaderProps, ITaxonomyPickerLoaderState> {
  constructor(props: ITaxonomyPickerLoaderProps, state: ITaxonomyPickerLoaderState) {
    super(props);

    this.state = {
      loadingScripts: true,
      errors: []
    };
  }

  public componentDidMount(): void {
    if (Environment.type === EnvironmentType.SharePoint
      || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._loadSPJSOMScripts();
    } else {
      this.setState({ loadingScripts: false, errors: [...this.state.errors, "You are on localhost mode (EnvironmentType.Local), be sure you disable termSetGuid and enable defaultOptions configuration in PropertyPaneTaxonomyPicker."] });
    }
  }

  public render(): JSX.Element {
    return (
      <div>
        {this.state.loadingScripts === false ?
          <TaxonomyPicker
            name="Department"
            placeholder={strings.TypeHerePlaceholder}
            displayName={strings.OwningDepartmentDisplayName}
            termSetGuid="d7b95001-2bd0-4aae-ad1d-9a6ed6b44061"
            termSetName="Department"
            termSetCountMaxSwapToAsync={100}
            multi={false}
            showPath
            isLoading={this.state.loadingScripts}
            onPickerChange={this.props.onPickerChange}
            defaultValue={this.props.defaultValue}
          />
          :
          null
        }
        {this.state.errors.length > 0 ? this.renderErrorMessage() : null}
      </div>
    );
  }

  private _loadSPJSOMScripts() {
    const siteColUrl = this.props.context.pageContext.web.absoluteUrl;
    try {
      SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/init.js', {
        globalExportsName: '$_global_init'
      })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/MicrosoftAjax.js', {
            globalExportsName: 'Sys'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.Runtime.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.taxonomy.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): void => {
          this.setState({ loadingScripts: false });
        })
        .catch((reason: any) => {
          this.setState({ loadingScripts: false, errors: [...this.state.errors, reason] });
        });
    } catch (error) {
      this.setState({ loadingScripts: false, errors: [...this.state.errors, error] });
    }
  }

  private renderErrorMessage() {
    return (
      <div>
        {this.state.errors}
      </div>
    );
  }
}