import * as React from 'react';
import * as strings from 'SiteCatalogWebPartStrings';
import styles from '../SiteCatalog.module.scss';
import { GraphHttpClient, HttpClientResponse, IGraphHttpClientOptions } from '@microsoft/sp-http';
import { IUserProfile } from '../interfaces/IUserProfile';
import { UserProfileService } from './UserProfileService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IUserProfileProps {
    context: WebPartContext;
    description?: string;
    userLoginName: string;
}

export interface IUserProfileState {
  firstName?: string;
  lastname?: string;
  userProfileProperties?: any[];
  isFirstName?: boolean;
  isLastName?: boolean;
  email?: string;
  isWorkPhone?: boolean;
  isDepartment?: boolean;
  displayName?: string;
  pictureUrl?: string;
  workPhone?: string;
  department?: string;
  isPictureUrl?: boolean;
  title?: string;
  office?: string;
  isOffice?: boolean;
  GUID?: string;
  isGUID?: boolean;
}

export default class UserProfile extends React.Component<IUserProfileProps, IUserProfileState> {

  constructor(props) {
    super(props);

    this.state = {
      firstName: "",
      lastname: "",
      userProfileProperties: [],
      isFirstName: false,
      isLastName: false,
      email: "",
      workPhone: "",
      department: "",
      pictureUrl: "",
      isPictureUrl: false,
      title: "",
      office: "",
      isOffice: false,
      GUID: "",
      isGUID: false
    };
  }

  public componentDidMount(){
    // Fetch data from graph
    this._getProperties();
  }

  private _getProperties(): void {
 
    const userProfileService: UserProfileService = new UserProfileService(this.props.context, '');
 
    userProfileService.getUserProfileProperties().then((response) => {
 
      this.setState({ userProfileProperties: response.UserProfileProperties });
      this.setState({ email: response.Email });
      this.setState({ displayName: response.DisplayName });
      this.setState({ title: response.Title });
 
      for (let i: number = 0; i < this.state.userProfileProperties.length; i++) {
 
        if (this.state.isFirstName == false || this.state.isLastName == false || this.state.isDepartment == false || this.state.isWorkPhone == false || this.state.isPictureUrl == false || this.state.isOffice == false || this.state.isGUID) {
 
          if (this.state.userProfileProperties[i].Key == "FirstName") {
            //this.state.isFirstName = true;
            this.setState({ isFirstName: true, firstName: this.state.userProfileProperties[i].Value });
          }
          if (this.state.userProfileProperties[i].Key == "LastName") {
            //this.state.isLastName = true;
            this.setState({ isLastName: true, lastname: this.state.userProfileProperties[i].Value });
          }
          if (this.state.userProfileProperties[i].Key == "WorkPhone") {
            //this.state.isWorkPhone = true;
            this.setState({ isWorkPhone: true, workPhone: this.state.userProfileProperties[i].Value });
          }
          if (this.state.userProfileProperties[i].Key == "Department") {
            //this.state.isDepartment = true;
            this.setState({ isDepartment: true, department: this.state.userProfileProperties[i].Value });
          }
          if (this.state.userProfileProperties[i].Key == "Office") {
            //this.state.isOffice = true;
            this.setState({ isOffice: true, office: this.state.userProfileProperties[i].Value });
          }
          if (this.state.userProfileProperties[i].Key == "PictureURL") {
            //this.state.isPictureUrl = true;
            this.setState({ isPictureUrl: true, pictureUrl: this.state.userProfileProperties[i].Value });
          }
          if (this.state.userProfileProperties[i].Key == "UserProfile_GUID"){
            this.setState({ isGUID:  true, GUID: this.state.userProfileProperties[i].Value });
          }
 
        }
 
      }
 
 
    });
 
  }

  public render(): React.ReactElement<IUserProfileProps> {

    return (
      <div>
        {this.state.displayName !== '' ? this.state.displayName : ''}
      </div>
    );
  }

}
