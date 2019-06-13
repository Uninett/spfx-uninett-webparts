import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IUserDirectoryProps {
  context: WebPartContext;
  api: string;
  showPhoto: boolean;
  showJobTitle: boolean;
  showMail: boolean;
  showPhone: boolean;
  compactMode: boolean;
}
