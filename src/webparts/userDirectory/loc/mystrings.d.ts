declare interface IUserDirectoryWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ApiLabel: string;
  CompactModeLabel: string;
  SearchFor: string;
  SearchForValidationErrorMessage: string;
}

declare module 'UserDirectoryWebPartStrings' {
  const strings: IUserDirectoryWebPartStrings;
  export = strings;
}
