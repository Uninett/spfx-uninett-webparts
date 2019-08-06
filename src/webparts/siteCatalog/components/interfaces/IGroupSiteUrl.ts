export interface IGroupSiteUrl {
    createdDateTime: string;
    description: string;
    id: string;
    lastModifiedDateTime: string;
    name: string;
    webUrl: string;
    displayName: string;
    siteCollection: { ISiteCollection };
}

export interface ISiteCollection {
    hostName: string;
}