import { SPHttpClient } from '@microsoft/sp-http';
import { SharePointUserPersona } from '../models/PeoplePicker';

export interface IPeoplePickerProps {
  description: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  typePicker: string;
  principalTypeUser: boolean;
  principalTypeSharePointGroup: boolean;
  principalTypeSecurityGroup: boolean;
  principalTypeDistributionList: boolean;
  numberOfItems: number;
  onChange?: (items: SharePointUserPersona[]) => void;
}
