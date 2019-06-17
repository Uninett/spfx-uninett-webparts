import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IUserDirectoryProps {
  context: WebPartContext;
  api: string;
  compactMode: boolean;
  alternatingColours: boolean;
  showPhoto: boolean;
  showJobTitle: boolean;
  showDepartment: boolean;
  showOfficeLocation: boolean;
  showCity: boolean; 
  showPhone: boolean;
  showMail: boolean;  
  colNameTitle: string;
  colJobTitleTitle: string;
  colDepartmentTitle: string;
  colOfficeLocationTitle: string;
  colCityTitle: string;
  colPhoneTitle: string;
  colMailTitle: string;
}
