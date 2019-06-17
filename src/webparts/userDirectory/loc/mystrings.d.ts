declare interface IUserDirectoryWebPartStrings {
  // Pages  
  PageGeneralDescription: string;
  PageColumnsDescription: string;
  // Groups
  GroupDataSource: string;
  GroupAppearance: string;
  GroupColumns: string;
  GroupColumnTitles: string;
  // Labels
  ApiLabel: string;
  CompactModeLabel: string;
  CompactModeOn: string;
  CompactModeOff: string;
  AlternateColoursLabel: string;
  AlternateColoursOn: string;
  AlternateColoursOff: string;
  // Column names
  ColPhotoText: string;
  ColNameText: string;
  ColJobTitleText: string;
  ColDepartmentText: string;
  ColOfficeLocationText: string;
  ColCityText: string;
  ColPhoneText: string;
  ColMailText: string;
  // Web part contents
  SearchBoxLabel: string;
}

declare module 'UserDirectoryWebPartStrings' {
  const strings: IUserDirectoryWebPartStrings;
  export = strings;
}
