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
  ColPhoto: string;
  ColName: string;
  ColJobTitle: string;
  ColDepartment: string;
  ColOfficeLocation: string;
  ColCity: string;
  ColPhone: string;
  ColMail: string;
  // Web part contents
  SearchBoxLabel: string;
}

declare module 'UserDirectoryWebPartStrings' {
  const strings: IUserDirectoryWebPartStrings;
  export = strings;
}
