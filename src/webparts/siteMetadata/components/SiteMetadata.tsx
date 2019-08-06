import * as React from 'react';
import styles from './SiteMetadata.module.scss';
import { ISiteMetadataProps } from './ISiteMetadataProps';
import { ISiteMetadataState } from './ISiteMetadataState';
import { Activity } from './Activity';
import { Department } from './Department';
import { Project } from './Project';
import { Group } from './Group';
import * as strings from 'SiteMetadataWebPartStrings';
import { UserProfileService } from './services/UserProfileService';
import { AadHttpClient, HttpClientResponse, IAadHttpClientOptions } from '@microsoft/sp-http';

export default class SiteMetadata extends React.Component<ISiteMetadataProps, ISiteMetadataState> {

  constructor(props) {
    super(props);
    this.state = {
      listData: null,
      personObject: null,
      displayNameField: "",
      parentDepartmentField: "",
      groupType: "",
      siteType: null,
      isLoading: true,
      errorMessage: "",
      hideDialog: true
    };
  }

  public componentWillMount() {
    // Fetch data from graph
    this.loadSiteList();
  }

  public render(): React.ReactElement<ISiteMetadataProps> {
    
    let view;
    let siteType = "";

    if (this.state.siteType) {

      switch (this.state.siteType.type) {
        case "aktivitet": {
          view = (<Activity
            context={this.props.context}
            displayMode={this.props.displayMode}
            orderSiteURL={this.props.orderSiteURL}
            editName={this.props.editName}
            editDescription={this.props.editDescription}
            editOwner={this.props.editOwner}
            editParentDepartment={this.props.editParentDepartment}
          />);
          siteType = strings.siteTypeActivity;
          break;
        }
        case "avdeling": {
          view = (<Department
            context={this.props.context}
            displayMode={this.props.displayMode}
            orderSiteURL={this.props.orderSiteURL}
            editName={this.props.editName}
            editDescription={this.props.editDescription}
            editOwner={this.props.editOwner}
            editShortName={this.props.editShortName}
            editParentDepartment={this.props.editParentDepartment}
          />);
          siteType = strings.siteTypeDepartment;
          break;
        }
        case "prosjekt": {
          view = (<Project
            context={this.props.context}
            displayMode={this.props.displayMode}
            orderSiteURL={this.props.orderSiteURL}
            editName={this.props.editName}
            editProjectGoal={this.props.editProjectGoal}
            editProjectPurpose={this.props.editProjectPurpose}
            editOwner={this.props.editOwner}
            editParentDepartment={this.props.editParentDepartment}
            // editOwningDepartment={this.props.editOwningDepartment}
            editProjectNumber={this.props.editProjectNumber}
            editStartDate={this.props.editStartDate}
            editEndDate={this.props.editEndDate}
          />);

          siteType = strings.siteTypeProject;
          break;
        }
        case "seksjon": {
          view = (<Department
            context={this.props.context}
            displayMode={this.props.displayMode}
            orderSiteURL={this.props.orderSiteURL}
            editName={this.props.editName}
            editDescription={this.props.editDescription}
            editOwner={this.props.editOwner}
            editShortName={this.props.editShortName}
            editParentDepartment={this.props.editParentDepartment}
          />);
          siteType = strings.siteTypeSection;
          break;
        }
        case "gruppe": {
          view = (<Group
            context={this.props.context}
            displayMode={this.props.displayMode}
            orderSiteURL={this.props.orderSiteURL}
            editName={this.props.editName}
          />);
          siteType = strings.siteTypeGroup;
          break;
        }

      }

    }

    return (
      <div className={styles.siteMetadata}>
        <div>
          <div>
            <div>
              {this.state.siteType ? <p className={'ms-font-xxl ' + styles.title}>{strings.informationAbout} {siteType}</p> : ''}
              {this.state.errorMessage ? 'Error: ' + this.state.errorMessage : ''}
              {view}
            </div>
          </div>
        </div>
      </div>
    );
  }

  private getSiteType = (groupType: string) => {
    let siteType: SiteType = { type: 'gruppe', label: 'gruppen' };

    if (groupType.startsWith('Aktivitet')) siteType = { type: 'aktivitet', label: 'aktiviteten' };
    else if (groupType.startsWith('Seksjon')) siteType = { type: 'seksjon', label: 'seksjonen' };
    else if (groupType.startsWith('Avdeling')) siteType = { type: 'avdeling', label: 'avdelingen' };
    else if (groupType.startsWith('Prosjekt')) siteType = { type: 'prosjekt', label: 'prosjektet' };

    return siteType;
  }

  public loadSiteList = () => {
    let handler = this;
    // Query for all groups on the tenant using Microsoft Graph.
    let groupId = this.props.context.pageContext.legacyPageContext.groupId;

    if (!groupId) {
      this.setState({
        errorMessage: 'This webpart can only be used on a group site.'
      });
      return;
    }

    this.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient) => {
        return client
          .get(
            `https://graph.microsoft.com/v1.0/groups/${groupId}?$select=displayName,id,extvcs569it_InmetaGenericSchema`,
            AadHttpClient.configurations.v1
          );
      })
      .then((response: HttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          console.warn(response.statusText);
        }
      })
      .then((result: any) => {
        let personObject;
        let userProfileService: UserProfileService;
        let siteType = this.getSiteType(result['extvcs569it_InmetaGenericSchema']['ValueString00']);
        userProfileService = new UserProfileService(this.props.context, result.extvcs569it_InmetaGenericSchema.ValueString03);
        userProfileService.getUserProfileProperties().then((userResult) => {
          personObject = { imageShouldFadeIn: true, imageUrl: "/_layouts/15/userphoto.aspx?size=S&accountname=" + userResult.Email, primaryText: userResult.DisplayName, secondaryText: "", selected: true, tertiaryText: "" };
          this.setState({
            listData: result,
            personObject: personObject,
            displayNameField: result['extvcs569it_InmetaGenericSchema']['ValueString01'],
            parentDepartmentField: result['extvcs569it_InmetaGenericSchema']['ValueString04'],
            groupType: result['extvcs569it_InmetaGenericSchema']['ValueString00'],
            siteType: siteType
          });
        });
      });
  }

}
