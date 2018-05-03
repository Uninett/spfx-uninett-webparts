declare interface ISocialLinksWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  facebookUrl: string;
  twitterUrl: string;
  linkedinUrl: string;
  youtubeUrl: string;
}

declare module 'SocialLinksWebPartStrings' {
  const strings: ISocialLinksWebPartStrings;
  export = strings;
}
