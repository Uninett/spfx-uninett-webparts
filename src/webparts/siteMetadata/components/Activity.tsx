import * as React from 'react';
import styles from './SiteMetadata.module.scss';
import { IActivityState } from './IActivityState';
import { GraphHttpClient, HttpClientResponse, IGraphHttpClientOptions } from '@microsoft/sp-http';
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
import { ParentDepartment } from './ParentDepartment';
import * as strings from 'SiteMetadataWebPartStrings';

export interface IActivityProps {
  displayMode: DisplayMode;
  context: WebPartContext;
  orderSiteURL: string;
  editName: boolean;
  editDescription: boolean;
  editOwner: boolean;
  editParentDepartment: boolean;
}

class Activity extends React.Component<IActivityProps, IActivityState> {

  constructor(props: any) {
    super(props);
    this.state = {
      listData: null,
      personObject: null,
      errorMessage: "",
      displayNameField: "",
      descriptionField: "",
      ownersField: "",
      parentDepartmentField: "",
      hideDialog: true
    };
  }

  componentWillMount() {
    // Fetch data from graph
    this.loadSiteList();
  }

  render() {
    return (<div>
      {this.state.listData ?
        this.props.displayMode === DisplayMode.Read
          ? this.renderDisplayMetadata()
          : this.renderEditMetadata()
        : ''}
    </div>)
  }

  private renderDisplayMetadata = () => {
    return this.state.listData !== null && this.state.listData.hasOwnProperty('extvcs569it_InmetaGenericSchema') ? (
      <div className={"ms-Grid " + styles.grid}>
        <div className={styles.row}>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.ActivityName}:</span> {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString01}</div>
          </div>
        </div>

        <div className={styles.row}>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.ActivityDescription}:</span> {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString02}</div>
          </div>
        </div>
        <div className={styles.row}>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.ActivityResponsible}:</span> {this.state.personObject.hasOwnProperty("primaryText") ? this.state.personObject.primaryText : this.state.listData.extvcs569it_InmetaGenericSchema.ValueString03}</div>
          </div>
        </div>

        {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString04 ? (
          <div className={styles.row}>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.ParentDepartment}:</span> {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString04}</div>
            </div>
          </div>
        ) : ''}
        

      </div>
    ) : ''
  }

  private renderEditMetadata = () => {
    return this.state.listData !== null && this.state.listData.hasOwnProperty('extvcs569it_InmetaGenericSchema') ? (
      <div className={"ms-Grid" + styles.siteMetadata}>

        <div className={styles.property}>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
              {this.props.editName ?
                <TextField maxLength={255} label={strings.ActivityName} defaultValue={this.state.listData.extvcs569it_InmetaGenericSchema.ValueString01} onChanged={(value: string) => { this.setState({ displayNameField: value }); }} />
                : ''}
            </div>
          </div>
        </div>

        <div className={styles.property}>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
              {this.props.editDescription ?
                <TextField maxLength={500} multiline rows={4} label={strings.ActivityDescription} defaultValue={this.state.listData.extvcs569it_InmetaGenericSchema.ValueString02} onChanged={(value: string) => { this.setState({ descriptionField: value }); }} />
                : ''}
            </div>
          </div>
        </div>

        <div className={styles.property}>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
              {this.props.editOwner ?
                <Label>{strings.ActivityResponsible}</Label>
                : ''}
            </div>
          </div>

          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
              {this.props.editOwner ?
                <PeoplePicker
                  description="Office UI Fabric People Picker"
                  spHttpClient={this.props.context.spHttpClient}
                  siteUrl={this.props.context.pageContext.site.absoluteUrl}
                  typePicker="Normal"
                  principalTypeUser={true}
                  principalTypeSharePointGroup={true}
                  principalTypeSecurityGroup={false}
                  principalTypeDistributionList={false}
                  numberOfItems={10}
                  defaultSelectedItems={[this.state.personObject]}
                  onChange={(people: SharePointUserPersona[]) => {
                    this.setState({ people })
                    var emails = people.map(spPersona => {
                      return spPersona.User.Email
                    })
                  }}
                />
                : ''}
            </div>
          </div>
        </div>

        <div className={styles.property}>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
              {this.props.editParentDepartment && this.state.listData.extvcs569it_InmetaGenericSchema.ValueString04 ?
                <ParentDepartment
                  context={this.props.context}
                  onChanged={(option) => {
                    this.setState({ parentDepartmentField: option.text })
                  }}
                  defaultSelectedKey={this.state.listData.extvcs569it_InmetaGenericSchema.ValueString04}
                  orderSiteURL={this.props.orderSiteURL}
                />
                : ''}
            </div>
          </div>

        </div>

        <div className={styles.property}>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
              <DefaultButton
                primary={true}
                data-automation-id='test'
                text={strings.Save}
                onClick={this.saveSiteList}
              />
            </div>
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

  loadSiteList = () => {
    let handler = this;
    // Query for all groupos on the tenant using Microsoft Graph.
    let groupId = this.props.context.pageContext.legacyPageContext.groupId;

    if (!groupId) {
      this.setState({
        errorMessage: 'This webpart can only be used on a group site.'
      });
      return;
    }

    this.props.context.graphHttpClient.get(`v1.0/groups/${groupId}?$select=displayName,id,extvcs569it_InmetaGenericSchema`, GraphHttpClient.configurations.v1).then((response: HttpClientResponse) => {
      if (response.ok) {
        return response.json();
      } else {
        console.warn(response.statusText);
      }
    }).then((result: any) => {
      let personObject;
      let userProfileService: UserProfileService;
      userProfileService = new UserProfileService(this.props.context, result.extvcs569it_InmetaGenericSchema.ValueString03);
      userProfileService.getUserProfileProperties().then((userResult) => {
        
        personObject = { imageShouldFadeIn: true, imageUrl: "/_layouts/15/userphoto.aspx?size=S&accountname=" + userResult.Email, primaryText: userResult.DisplayName, secondaryText: "", selected: true, tertiaryText: "" }
        this.setState({
          listData: result,
          personObject: personObject,
          displayNameField: result['extvcs569it_InmetaGenericSchema']['ValueString01'],
          parentDepartmentField: result['extvcs569it_InmetaGenericSchema']['ValueString04']
        })
      });
    });
  }

  saveSiteList = () => {
    let handler = this;
    let data = this.buildMetaData();
    // Query for all groupos on the tenant using Microsoft Graph.
    let groupId = this.props.context.pageContext.legacyPageContext.groupId;
    this.props.context.graphHttpClient.fetch(`v1.0/groups/${groupId}`, GraphHttpClient.configurations.v1, {
      method: "PATCH",
      body: JSON.stringify(data)
    }).then((response: HttpClientResponse) => {
      if (response.ok) {
        // Success!
        handler.setState({ hideDialog: false })
        handler.loadSiteList();
      } else {
        console.warn(response.statusText);
      }
    })
  }

  buildMetaData = () => {
    if (this.state.personObject != [] && this.state.displayNameField != "") {

      if (this.state.listData.extvcs569it_InmetaGenericSchema.ValueString04 && this.state.parentDepartmentField == "") {
        alert(strings.IncompleteAlert);
      } else {
        let data = {
          "extvcs569it_InmetaGenericSchema": {}
        };
        if (this.state.displayNameField) {
          data["extvcs569it_InmetaGenericSchema"]["ValueString01"] = this.state.displayNameField;
          data["extvcs569it_InmetaGenericSchema"]["LabelString01"] = this.buildDisplayNameLabel(this.state.displayNameField);
        }
        if (this.state.descriptionField) {
          data["extvcs569it_InmetaGenericSchema"]["ValueString02"] = this.state.descriptionField;
        }
        if (this.state.people && this.state.people.length !== 0) {
          data["extvcs569it_InmetaGenericSchema"]["ValueString03"] = this.state.people[0].User.Description;
        }
        if (this.state.parentDepartmentField) {
          data["extvcs569it_InmetaGenericSchema"]["ValueString04"] = this.state.parentDepartmentField;
        }
  
        return data;
      }
      
    }
    else {
      alert(strings.IncompleteAlert);
    }
  }

  buildDisplayNameLabel = (displayName:string):string => {
    let prefix = 'Aktivitet: ';

    if (!displayName.includes(prefix)) {
      return displayName;
    }

    let label = displayName.replace(prefix, '');

    return label;
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

}

export { Activity }