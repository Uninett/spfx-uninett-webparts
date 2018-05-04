
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLink
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SocialLinksWebPart.module.scss';
import * as strings from 'SocialLinksWebPartStrings';

export interface ISocialLinksWebPartProps {
  facebookField: string;
  twitterField: string;
  linkedinField: string;
  youtubeField: string;
  bgColorField: string;
}


// This line is required for the image path
//require('set-webpack-public-path!')
require("@microsoft/loader-set-webpack-public-path!");

export default class SocialLinksWebPart extends BaseClientSideWebPart<ISocialLinksWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${ styles.socialLinks }">
      <div class="${ styles.container }">
        <div class="${ styles.row }" style="background-color:${escape(this.properties.bgColorField)};">
          <div class="${ styles.column }">

          

        <div class="${ styles.socialContent}">
          <table>
            <tbody>
              <tr>
                <td><a href="${escape(this.properties.facebookField)}"><img src="${require<string>('./images/facebook_white.svg')}" class="svg"></a></td>
                <td><a href="${escape(this.properties.twitterField)}"><img src="${require<string>('./images/twitter_white.svg')}" class="svg"></a></td>
                <td><a href="${escape(this.properties.linkedinField)}"><img src="${require<string>('./images/linkedin_white.svg')}" class="svg"></a></td>
                <td><a href="${escape(this.properties.youtubeField)}"><img src="${require<string>('./images/youtube_white.svg')}" class="svg"></a></td>
              </tr>
            </tbody>
          </table>
        </div>
          </div>
        </div> 
      </div>
    </div>`;
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
                PropertyPaneTextField('facebookField', {
                  label: strings.facebookUrl,
                  value: this.properties.facebookField
                }),
                PropertyPaneTextField('twitterField', {
                  label: strings.twitterUrl,
                  value: this.properties.twitterField
                }),
                PropertyPaneTextField('linkedinField', {
                  label: strings.linkedinUrl,
                  value: this.properties.linkedinField,
                }),
                PropertyPaneTextField('youtubeField', {
                  label: strings.youtubeUrl,
                  value: this.properties.youtubeField
                })
              ]
            },
            {
              groupName: strings.ColorGroupName,
              groupFields: [
                PropertyPaneTextField('bgColorField', {
                  label: strings.bgColorLabel,
                  value: this.properties.bgColorField
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
