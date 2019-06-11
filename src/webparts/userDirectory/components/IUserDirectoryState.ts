import { IUserItem } from './IUserItem';

export interface IUserDirectoryState {
  users: Array<IUserItem>;
  searchFor: string;
}