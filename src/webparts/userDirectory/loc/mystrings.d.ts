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
  BtnResetText: string;
  // Column names
  ColPhotoTitle: string;
  ColNameTitle: string;
  ColJobTitleTitle: string;
  ColDepartmentTitle: string;
  ColOfficeLocationTitle: string;
  ColCityTitle: string;
  ColPhoneTitle: string;
  ColMailTitle: string;
  // Web part contents
  SearchBoxLabel: string;
}

declare module 'UserDirectoryWebPartStrings' {
  const strings: IUserDirectoryWebPartStrings;
  export = strings;
}
