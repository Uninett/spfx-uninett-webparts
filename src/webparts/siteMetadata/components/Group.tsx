import * as React from 'react';
import styles from './SiteMetadata.module.scss';
import { IGroupState } from './IGroupState';
import { AadHttpClient, HttpClientResponse, IAadHttpClientOptions } from '@microsoft/sp-http';
import { DisplayMode } from "@microsoft/sp-core-library";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { IWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IClientPeoplePickerSearchUser,
  IEnsurableSharePointUser,
  IEnsureUser,
  IPeoplePickerState,
  SharePointUserPersona
}
  from '../components/models/PeoplePicker';
import { PeoplePicker } from './PeoplePicker';
import { UserProfileService } from './services/UserProfileService';
import { Label } from 'office-ui-fabric-react';
import * as strings from 'SiteMetadataWebPartStrings';

export interface IGroupProps {
  displayMode: DisplayMode;
  context: WebPartContext;
  orderSiteURL: string;
  editName: boolean;
}

class Group extends React.Component<IGroupProps, IGroupState> {

  constructor(props: any) {
    super(props);
    this.state = {
      listData: null,
      personObject: null,
      errorMessage: "",
      displayNameField: "",
      ownersField: "",
      hideDialog: true
    };
  }

  public componentWillMount() {
    // Fetch data from graph
    this.loadSiteList();
  }

  public render() {
    return (<div>
      {this.state.listData ?
        this.props.displayMode === DisplayMode.Read
          ? this.renderDisplayMetadata()
          : this.renderEditMetadata()
        : ''}
    </div>);
  }

  private renderDisplayMetadata = () => {
    return this.state.listData !== null && this.state.listData.hasOwnProperty('extvcs569it_InmetaGenericSchema') ? (
      <div className={"ms-Grid " + styles.grid}>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.DepartmentName}:</span> {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString01}</div>
        </div>
      </div>
    ) : '';
  }

  private renderEditMetadata = () => {
    return this.state.listData !== null && this.state.listData.hasOwnProperty('extvcs569it_InmetaGenericSchema') ? (
      <div className="ms-Grid">

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            {this.props.editName ?
              <TextField maxLength={255} label={strings.DepartmentName} defaultValue={this.state.listData.extvcs569it_InmetaGenericSchema.ValueString01} onChange={(_, value: string) => { this.setState({ displayNameField: value }); }} />
              : ''}
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            <DefaultButton
              primary={true}
              data-automation-id='test'
              text='Lagre'
              onClick={this.saveSiteList}
            />
          </div>
        </div>

        <Dialog
          hidden={this.state.hideDialog}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: strings.GroupSavedSuccesfullyDialog,
            subText: strings.GroupSavedSuccesfullyDialogSub
          }}
          modalProps={{
            titleAriaId: 'myLabelId',
            subtitleAriaId: 'mySubTextId',
            isBlocking: false,
            containerClassName: 'ms-dialogMainOverride'
          }}
        >
          {null /** You can also include null values as the result of conditionals */}
          <DialogFooter>
            <DefaultButton onClick={this._closeDialog} text={strings.ok} />
          </DialogFooter>
        </Dialog>
      </div>
    ) : '';
  }

  public loadSiteList = () => {
    let handler = this;
    // Query for all groupos on the tenant using Microsoft Graph.
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
        userProfileService = new UserProfileService(this.props.context, result.extvcs569it_InmetaGenericSchema.ValueString03);
        userProfileService.getUserProfileProperties().then((userResult) => {
          personObject = { imageShouldFadeIn: true, imageUrl: "/_layouts/15/userphoto.aspx?size=S&accountname=" + userResult.Email, primaryText: userResult.DisplayName, secondaryText: "", selected: true, tertiaryText: "" };
          this.setState({
            listData: result,
            displayNameField: result['extvcs569it_InmetaGenericSchema']['ValueString01'],
          });
        });
      });
  }

  public saveSiteList = () => {
    let handler = this;
    let data = this.buildMetaData();
    // Query for all groupos on the tenant using Microsoft Graph.
    let groupId = this.props.context.pageContext.legacyPageContext.groupId;
    this.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient) => {
        return client
          .fetch(
            `https://graph.microsoft.com/v1.0/groups/${groupId}`,
            AadHttpClient.configurations.v1, 
            {
              method: "PATCH",
              body: JSON.stringify(data)
            }
          );
      })
      .then((response: HttpClientResponse) => {
        if (response.ok) {
          // Success!
          handler.setState({ hideDialog: false });
          handler.loadSiteList();
        } else {
          console.warn(response.statusText);
        }
      });
  }

  public buildMetaData = () => {
    if (this.state.displayNameField != "") {

      if (this.state.listData.extvcs569it_InmetaGenericSchema.ValueString04 == "") {
        alert(strings.IncompleteAlert);
      } else {
        let data = {
          "extvcs569it_InmetaGenericSchema": {}
        };
        if (this.state.displayNameField) {
          data["extvcs569it_InmetaGenericSchema"]["ValueString01"] = this.state.displayNameField;
          data["extvcs569it_InmetaGenericSchema"]["LabelString01"] = this.buildDisplayNameLabel(this.state.displayNameField);
        }
  
        return data;
      }
      
    }
    else {
      alert(strings.IncompleteAlert);
    }
  }

  public buildDisplayNameLabel = (displayName:string):string => {
    let prefix = 'Avdeling: ';
    let prefix2 = 'Seksjon: ';

    if (!displayName.includes(prefix) && !displayName.includes(prefix2)) {
      return displayName;
    }

    let label = displayName.replace(prefix, '');
    label = displayName.replace(prefix2, '');
    
    return label;
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

}

export { Group };