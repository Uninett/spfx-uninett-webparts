declare interface ISocialLinksWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ColorGroupName: string;
  facebookUrl: string;
  twitterUrl: string;
  linkedinUrl: string;
  youtubeUrl: string;
  bgColorLabel: string;
  iconColorLabel: string;
  iconColorOffLabel: string;
  iconColorOnLabel: string;
}

declare module 'SocialLinksWebPartStrings' {
  const strings: ISocialLinksWebPartStrings;
  export = strings;
}
