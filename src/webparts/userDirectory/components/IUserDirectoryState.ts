import { IUserItem } from './IUserItem';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';

export interface IUserDirectoryState {
  users: Array<IUserItem>;
  columns: IColumn[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
}