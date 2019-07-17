import * as React from 'react';
import styles from './SiteMetadata.module.scss';
import { IProjectState } from './IProjectState';
import { GraphHttpClient, HttpClientResponse, IGraphHttpClientOptions } from '@microsoft/sp-http';
import { DisplayMode } from "@microsoft/sp-core-library";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { IWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import * as strings from 'SiteMetadataWebPartStrings';
import { autobind } from 'office-ui-fabric-react';
import { PeoplePicker } from './PeoplePicker';
import { ParentDepartment } from './ParentDepartment';
import {
  IClientPeoplePickerSearchUser,
  IEnsurableSharePointUser,
  IEnsureUser,
  IPeoplePickerState,
  SharePointUserPersona
}
  from '../components/models/PeoplePicker';
import { IPersonaProps, Label } from 'office-ui-fabric-react';
import { UserProfileService } from './services/UserProfileService';
import { ITaxonomyObject } from './interfaces/ITaxonomyObject';
import TaxonomyPickerLoader from './TaxonomyPicker/TaxonomyPickerLoader'
import "react-taxonomypicker/dist/React.TaxonomyPicker.css";


export interface IProjectProps {
  displayMode: DisplayMode;
  context: WebPartContext;
  orderSiteURL: string;
  editName: boolean;
  editProjectGoal: boolean;
  editProjectPurpose: boolean;
  editOwner: boolean;
  editParentDepartment: boolean;
  //editOwningDepartment: boolean;
  editProjectNumber: boolean;
  editStartDate: boolean;
  editEndDate: boolean;
}

const DayPickerStrings: IDatePickerStrings = {
  months: [
    strings.January,
    strings.February,
    strings.March,
    strings.April,
    strings.May,
    strings.June,
    strings.July,
    strings.August,
    strings.September,
    strings.October,
    strings.November,
    strings.December
  ],

  shortMonths: [
    strings.Jan,
    strings.Feb,
    strings.Mar,
    strings.Apr,
    strings.May,
    strings.Jun,
    strings.Jul,
    strings.Aug,
    strings.Sep,
    strings.Oct,
    strings.Nov,
    strings.Dec,
  ],

  days: [
    strings.Sunday,
    strings.Monday,
    strings.Tuesday,
    strings.Wednesday,
    strings.Thursday,
    strings.Friday,
    strings.Saturday
  ],

  shortDays: [
    strings.ShortMonday,
    strings.ShortTuesday,
    strings.ShortWednesday,
    strings.ShortThirsday,
    strings.ShortFriday,
    strings.ShortSaturday,
    strings.ShortSunday,
  ],

  goToToday: strings.goToToday,
  prevMonthAriaLabel: strings.prevMonthAriaLabel,
  nextMonthAriaLabel: strings.nextMonthAriaLabel,
  prevYearAriaLabel: strings.prevYearAriaLabel,
  nextYearAriaLabel: strings.nextYearAriaLabel
};

class Project extends React.Component<IProjectProps, IProjectState> {

  constructor(props: any) {
    super(props);
    this.state = {
      listData: null,
      personObject: null,
      errorMessage: "",
      displayNameField: "",
      ownersField: "",
      parentDepartmentField: "",
      projectNumberField: "",
      projectGoalField: null,
      projectPurposeField: null,
      hideDialog: true,
      people: []
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
    let taxonomyStringLabel
    if (this.state.listData.extvcs569it_InmetaGenericSchema.ValueString05) {
      let taxonomyString = this.state.listData.extvcs569it_InmetaGenericSchema.ValueString05;
      let taxonomyArray = taxonomyString.split("|");
      taxonomyStringLabel = taxonomyArray[0];
    }
    else {
      taxonomyStringLabel = "";
    }

    let startDateUnformated = this.state.listData.extvcs569it_InmetaGenericSchema.ValueDateTime00;
    let startDateArray = startDateUnformated.split("T");
    let startDateFormated = startDateArray[0];
    let endDateUnformated = this.state.listData.extvcs569it_InmetaGenericSchema.ValueDateTime01;
    let endDateArray = endDateUnformated.split("T");
    let endDateFormated = endDateArray[0];

    return this.state.listData !== null && this.state.listData.hasOwnProperty('extvcs569it_InmetaGenericSchema') ? (
      <div className={"ms-Grid " + styles.grid}>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.ProjectName}:</span> {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString01}</div>
        </div>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.ProjectGoal}:</span> {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString08}</div>
        </div>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.ProjectPurpose}:</span> {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString09}</div>
        </div>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.ProjectLeader}:</span>  {this.state.personObject.hasOwnProperty("primaryText") ? this.state.personObject.primaryText : this.state.listData.extvcs569it_InmetaGenericSchema.ValueString03}</div>
        </div>

        {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString04 ? (
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.ParentDepartment}:</span> {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString04}</div>
          </div>
        ) : ''}
        
        {/*<div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.OwningDepartmentDisplayName}:</span> {taxonomyStringLabel}</div>
    </div>*/}
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.ProjectNumber}:</span> {this.state.listData.extvcs569it_InmetaGenericSchema.ValueString06}</div>
        </div>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.StartDate}:</span> {startDateFormated}</div>
        </div>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"><span className={styles.span}>{strings.EndDate}:</span> {endDateFormated}</div>
        </div>
      </div>
    ) : ''
  }

  private renderEditMetadata = () => {
    let startDate = this.state.startDateField ? this.state.startDateField : new Date(this.state.listData.extvcs569it_InmetaGenericSchema.ValueDateTime00);
    let endDate = this.state.endDateField ? this.state.endDateField : new Date(this.state.listData.extvcs569it_InmetaGenericSchema.ValueDateTime01);

    // let taxonomyObject;
    // if (this.state.listData.extvcs569it_InmetaGenericSchema.ValueString05) {
    //   let taxonomyString = this.state.listData.extvcs569it_InmetaGenericSchema.ValueString05;
    //   let taxonomyArray = taxonomyString.split("|");
    //   let taxonomyStringLabel = taxonomyArray[0];
    //   let taxonomyStringID = taxonomyArray[1];
    //   taxonomyObject = { label: taxonomyStringLabel, path: taxonomyStringLabel, value: taxonomyStringID }
    // }
    // else {
    //   taxonomyObject = { label: "", path: "", value: "" };
    // }


    return this.state.listData !== null && this.state.listData.hasOwnProperty('extvcs569it_InmetaGenericSchema') ? (
      <div className="ms-Grid">

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            {this.props.editName ?
              <TextField maxLength={255} label={strings.ProjectName} defaultValue={this.state.listData.extvcs569it_InmetaGenericSchema.ValueString01} onChanged={(value: string) => { this.setState({ displayNameField: value }); }} />
              : ''}
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            {this.props.editProjectGoal ?
              <TextField maxLength={255} multiline rows={4} label={strings.ProjectGoal} defaultValue={this.state.listData.extvcs569it_InmetaGenericSchema.ValueString08} onChanged={(value: string) => { this.setState({ projectGoalField: value }); }} />
              : ''}
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            {this.props.editProjectPurpose ?
              <TextField maxLength={500} multiline rows={4} label={strings.ProjectPurpose} defaultValue={this.state.listData.extvcs569it_InmetaGenericSchema.ValueString09} onChanged={(value: string) => { this.setState({ projectPurposeField: value }); }} />
              : ''}
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            {this.props.editOwner ?
              <Label>{strings.ProjectLeader}</Label>
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

        
        {/* <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            {this.props.editOwningDepartment ?
              <TaxonomyPickerLoader
                context={this.props.context}
                multi
                name={strings.Department}
                onPickerChange={(Name, Option) => { this._onTaxonomyChanged(Name, Option) }}
                required={true}
                defaultValue={taxonomyObject}
              />
              : ''}
          </div>
        </div> */}

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            {this.props.editProjectNumber ?
              <TextField maxLength={20} label={strings.ProjectNumber} defaultValue={this.state.listData.extvcs569it_InmetaGenericSchema.ValueString06} onChanged={(value: string) => { this.setState({ projectNumberField: value }); }} />
              : ''}
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            {this.props.editStartDate ?
              <DatePicker label={strings.StartDate} strings={DayPickerStrings} showWeekNumbers={true} firstWeekOfYear={1} showMonthPickerAsOverlay={true} placeholder='Select a date...'
                value={startDate} onSelectDate={newDate => { this.setState({ startDateField: newDate }) }}
              />
              : ''}
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            {this.props.editEndDate ?
              <DatePicker label={strings.EndDate} strings={DayPickerStrings} showWeekNumbers={true} firstWeekOfYear={1} showMonthPickerAsOverlay={true} placeholder='Select a date...'
                value={endDate} minDate={startDate} onSelectDate={newDate => { this.setState({ endDateField: newDate }) }}
              />
              : ''}
          </div>
        </div>

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
        if (this.state.projectGoalField) {
          data["extvcs569it_InmetaGenericSchema"]["ValueString08"] = this.state.projectGoalField;
        }
        if (this.state.projectPurposeField) {
          data["extvcs569it_InmetaGenericSchema"]["ValueString09"] = this.state.projectPurposeField;
        }
        if (this.state.people && this.state.people.length !== 0) {
          data["extvcs569it_InmetaGenericSchema"]["ValueString03"] = this.state.people[0].User.Description;
        }
        if (this.state.parentDepartmentField) {
          data["extvcs569it_InmetaGenericSchema"]["ValueString04"] = this.state.parentDepartmentField;
        }
        // if (this.state.ownedDepartmentField) {
        //   data["extvcs569it_InmetaGenericSchema"]["ValueString05"] = this.state.ownedDepartmentField;
        // }
        if (this.state.projectNumberField) {
          data["extvcs569it_InmetaGenericSchema"]["ValueString06"] = this.state.projectNumberField;
        }
        if (this.state.startDateField) {
          data["extvcs569it_InmetaGenericSchema"]["ValueDateTime00"] = this.state.startDateField;
        }
        if (this.state.endDateField) {
          data["extvcs569it_InmetaGenericSchema"]["ValueDateTime01"] = this.state.endDateField;
        }
  
  
        return data;
      }

      
    }
    else {
      alert(strings.IncompleteAlert);
    }
  }

  buildDisplayNameLabel = (displayName:string):string => {
    let prefix = 'Prosjekt: ';

    if (!displayName.includes(prefix)) {
      return displayName;
    }

    let label = displayName.replace(prefix, '');
    
    return label;
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

  // @autobind
  // private _onTaxonomyChanged(Name, Option) {
  //   if (!Option) {
  //     return;
  //   }
  //   let TaxValue = "" + Option.label + "|" + Option.value + "";
  //   this.setState({ ownedDepartmentField: TaxValue });
  // }

}

export { Project }