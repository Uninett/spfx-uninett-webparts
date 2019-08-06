import { IUserProfile } from '../interfaces/IUserProfile';
import { WebPartContext } from '@microsoft/sp-webpart-base';
//import { IUserProfileProps } from './UserProfile';
import { SPHttpClient, HttpClientResponse, IAadHttpClientOptions } from '@microsoft/sp-http';
 
interface IUserProfileService {
  getUserProfileProperties: Promise<IUserProfile>;
  webAbsoluteUrl: string;
  userLoginName: string;
  context: WebPartContext;
}

export interface IUserProfileProps {
    context: WebPartContext;
    description?: string;
    userLoginName: string;
}
 
export class UserProfileService {
 
  private context: WebPartContext;
  private userLoginName: string;
 
  constructor(_context, _userLoginName) {
    //this.props = _props;
    this.context = _context;
    this.userLoginName = encodeURIComponent(_userLoginName);

    //this.encodedUserLoginName = this.props.userLoginName;
  }
 
  public getUserProfileProperties(): Promise<IUserProfile> {

      /*return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|${this.userLoginName}'`, SPHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        return response.json();
      });*/

      return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|${this.userLoginName}'`, SPHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        return response.json();
      });
  }
}