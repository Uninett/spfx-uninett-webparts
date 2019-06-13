import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IUserDirectoryProps {
  context: WebPartContext;
  api: string;
  compactMode: boolean;
}
