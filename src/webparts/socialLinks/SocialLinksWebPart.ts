
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLink,
  PropertyPaneToggle
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
  iconColorField: boolean;
}


// This line is required for the image path
//require('set-webpack-public-path!')
require("@microsoft/loader-set-webpack-public-path!");

export default class SocialLinksWebPart extends BaseClientSideWebPart<ISocialLinksWebPartProps> {



  public render(): void {
    var iconColor = "";

    if(this.properties.iconColorField == true) {
      iconColor = "white";
    }
    else {
      iconColor = "black";
    }

    var facebookSvg = require<string>("./images/facebook_" + iconColor + ".svg");
    var linkedinSvg = require<string>("./images/linkedin_" + iconColor + ".svg");
    var twitterSvg = require<string>("./images/twitter_" + iconColor + ".svg");
    var youtubeSvg = require<string>("./images/youtube_" + iconColor + ".svg");

    var dynHtmlIcons = "";
    
    if(this.properties.facebookField != "") {
      dynHtmlIcons += `<td><a href="${escape(this.properties.facebookField)}"><img src="${facebookSvg}"></a></td>`;
    }
    if(this.properties.twitterField != "") {
      dynHtmlIcons += `<td><a href="${escape(this.properties.twitterField)}"><img src="${twitterSvg}"></a></td>`;
    }
    if(this.properties.linkedinField != "") {
      dynHtmlIcons += `<td><a href="${escape(this.properties.linkedinField)}"><img src="${linkedinSvg}"></a></td>`;
    }
    if(this.properties.youtubeField != "") {
      dynHtmlIcons += `<td><a href="${escape(this.properties.youtubeField)}"><img src="${youtubeSvg}"></a></td>`;
    }

    var dynHtml = `
    <div class="${ styles.socialLinks }">
      <div class="${ styles.container }">
        <div class="${ styles.row }" style="background-color:${escape(this.properties.bgColorField)};">
          <div class="${ styles.column }">

            <div class="${ styles.socialContent}">
              <table>
                <tbody>
                  <tr>` + dynHtmlIcons + `</tr>
                </tbody>
              </table>
            </div>

          </div>
        </div> 
      </div>
    </div>`;

    this.domElement.innerHTML = dynHtml;
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
                }),
                PropertyPaneToggle('iconColorField', {
                  label: strings.iconColorLabel,
                  offText: strings.iconColorOffLabel,
                  onText: strings.iconColorOnLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
