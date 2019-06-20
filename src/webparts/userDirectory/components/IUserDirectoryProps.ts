import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserItem } from './IUserItem';

export interface IUserDirectoryProps {
  context: WebPartContext;
  api: string;
  isApiChanged: boolean;
  compactMode: boolean;
  alternatingColours: boolean;
  useBuiltInSearch: boolean;
  searchBoxPlaceholder: string;
  showPhoto: boolean;
  showName: boolean;
  showJobTitle: boolean;
  showDepartment: boolean;
  showOfficeLocation: boolean;
  showCity: boolean; 
  showPhone: boolean;
  showMail: boolean;  
  colName: string;
  colJobTitle: string;
  colDepartment: string;
  colOfficeLocation: string;
  colCity: string;
  colPhone: string;
  colMail: string;
}
