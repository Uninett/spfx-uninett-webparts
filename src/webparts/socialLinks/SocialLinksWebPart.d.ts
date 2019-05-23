import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
export interface ISocialLinksWebPartProps {
    facebookField: string;
    twitterField: string;
    linkedinField: string;
    youtubeField: string;
    bgColorField: string;
    iconColorField: boolean;
}
export default class SocialLinksWebPart extends BaseClientSideWebPart<ISocialLinksWebPartProps> {
    render(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
