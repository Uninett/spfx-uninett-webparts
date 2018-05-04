declare interface ISocialLinksWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ColorGroupName: string;
  facebookUrl: string;
  twitterUrl: string;
  linkedinUrl: string;
  youtubeUrl: string;
  bgColorLabel: string;
}

declare module 'SocialLinksWebPartStrings' {
  const strings: ISocialLinksWebPartStrings;
  export = strings;
}
