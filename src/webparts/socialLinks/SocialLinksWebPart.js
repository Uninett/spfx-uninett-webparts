"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var sp_lodash_subset_1 = require("@microsoft/sp-lodash-subset");
var SocialLinksWebPart_module_scss_1 = require("./SocialLinksWebPart.module.scss");
var strings = require("SocialLinksWebPartStrings");
// This line is required for the image path
//require('set-webpack-public-path!')
require("@microsoft/loader-set-webpack-public-path!");
var SocialLinksWebPart = /** @class */ (function (_super) {
    __extends(SocialLinksWebPart, _super);
    function SocialLinksWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    SocialLinksWebPart.prototype.render = function () {
        var iconColor = "";
        if (this.properties.iconColorField == true) {
            iconColor = "white";
        }
        else {
            iconColor = "black";
        }
        var facebookSvg = require("./images/facebook_" + iconColor + ".svg");
        var linkedinSvg = require("./images/linkedin_" + iconColor + ".svg");
        var twitterSvg = require("./images/twitter_" + iconColor + ".svg");
        var youtubeSvg = require("./images/youtube_" + iconColor + ".svg");
        var dynHtmlIcons = "";
        if (this.properties.facebookField != "") {
            dynHtmlIcons += "<td><a href=\"" + sp_lodash_subset_1.escape(this.properties.facebookField) + "\"><img src=\"" + facebookSvg + "\"></a></td>";
        }
        if (this.properties.twitterField != "") {
            dynHtmlIcons += "<td><a href=\"" + sp_lodash_subset_1.escape(this.properties.twitterField) + "\"><img src=\"" + twitterSvg + "\"></a></td>";
        }
        if (this.properties.linkedinField != "") {
            dynHtmlIcons += "<td><a href=\"" + sp_lodash_subset_1.escape(this.properties.linkedinField) + "\"><img src=\"" + linkedinSvg + "\"></a></td>";
        }
        if (this.properties.youtubeField != "") {
            dynHtmlIcons += "<td><a href=\"" + sp_lodash_subset_1.escape(this.properties.youtubeField) + "\"><img src=\"" + youtubeSvg + "\"></a></td>";
        }
        var dynHtml = "\n    <div class=\"" + SocialLinksWebPart_module_scss_1.default.socialLinks + "\">\n      <div class=\"" + SocialLinksWebPart_module_scss_1.default.container + "\">\n        <div class=\"" + SocialLinksWebPart_module_scss_1.default.row + "\" style=\"background-color:" + sp_lodash_subset_1.escape(this.properties.bgColorField) + ";\">\n          <div class=\"" + SocialLinksWebPart_module_scss_1.default.column + "\">\n\n            <div class=\"" + SocialLinksWebPart_module_scss_1.default.socialContent + "\">\n              <table>\n                <tbody>\n                  <tr>" + dynHtmlIcons + "</tr>\n                </tbody>\n              </table>\n            </div>\n\n          </div>\n        </div> \n      </div>\n    </div>";
        this.domElement.innerHTML = dynHtml;
    };
    Object.defineProperty(SocialLinksWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    SocialLinksWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                sp_webpart_base_1.PropertyPaneTextField('facebookField', {
                                    label: strings.facebookUrl,
                                    value: this.properties.facebookField
                                }),
                                sp_webpart_base_1.PropertyPaneTextField('twitterField', {
                                    label: strings.twitterUrl,
                                    value: this.properties.twitterField
                                }),
                                sp_webpart_base_1.PropertyPaneTextField('linkedinField', {
                                    label: strings.linkedinUrl,
                                    value: this.properties.linkedinField,
                                }),
                                sp_webpart_base_1.PropertyPaneTextField('youtubeField', {
                                    label: strings.youtubeUrl,
                                    value: this.properties.youtubeField
                                })
                            ]
                        },
                        {
                            groupName: strings.ColorGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('bgColorField', {
                                    label: strings.bgColorLabel,
                                    value: this.properties.bgColorField
                                }),
                                sp_webpart_base_1.PropertyPaneToggle('iconColorField', {
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
    };
    return SocialLinksWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = SocialLinksWebPart;
//# sourceMappingURL=SocialLinksWebPart.js.map