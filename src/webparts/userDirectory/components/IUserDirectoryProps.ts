import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IUserDirectoryProps {
  context: WebPartContext;
  api: string;
  showPhoto: boolean;
  showJobTitle: boolean;
  showDepartment: boolean;  
  showPhone: boolean;
  showMail: boolean;
  compactMode: boolean;
}
