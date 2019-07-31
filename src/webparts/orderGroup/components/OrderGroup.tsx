import * as React from 'react';
import * as $ from 'jquery';
import styles from './OrderGroup.module.scss';
import { IOrderGroupProps } from './IOrderGroupProps';
import { IOrderGroupState } from './IOrderGroupState';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Pivot, PivotItem, IPivotStyles } from 'office-ui-fabric-react/lib/Pivot';
import { GroupTypeChoice } from './GroupTypeChoice';
import { Activity } from './Activity';
import { Department } from './Department';
import { Project } from './Project';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';
import * as strings from 'OrderGroupWebPartStrings';
import { IStyleSet } from '@uifabric/styling';
require('./CustomStyles.module.scss');

const pivotStyles: Partial<IPivotStyles> = {
  root: { display: "none" }
};

export default class OrderGroup extends React.Component<IOrderGroupProps, IOrderGroupState> {

  constructor(props: any) {
    super(props);
    this.state = {
      showPanel: false,
      selectedKey: 0,
      hideDialog: true
    };
    this.handleGroupTypeChoice = this.handleGroupTypeChoice.bind(this);
    this.loadActivityPivot = this.loadActivityPivot.bind(this);
    this.loadDepartmentPivot = this.loadDepartmentPivot.bind(this);
    this.loadProjectPivot = this.loadProjectPivot.bind(this);
  }

  public handleGroupTypeChoice(groupType: string) {
    this.setState({ groupType: groupType });
    if (groupType == "Activity") {
      this.loadDepartmentPivot();
    }
    else if (groupType == "Department") {
      this.loadActivityPivot();
    }
    else {
      this.loadProjectPivot();
    }
  }

  public componentDidMount() {
    let  handler  =  this;
    (window  as  any).addEventListener('createSiteButtonClicked', (e)  =>  {

      handler.setState({
        showPanel:  !handler.state.showPanel
      });
    });
  }

  public render(): React.ReactElement<IOrderGroupProps> {
    
    return (
      <div>

        <div>
          <DefaultButton
            primary={true}
            data-automation-id='openPanel'
            text={strings.CreateSiteButton}
            onClick={this._showPanel}
          />
        </div>

        <Panel
          isOpen={this.state.showPanel}
          isLightDismiss={true}
          onDismiss={this._closePanel}
          type={PanelType.largeFixed}
          headerText={strings.CreateSiteTitle}
        >

          <div className={styles.orderGroup}>
            <div>
              <Pivot styles={ pivotStyles } selectedKey={`${this.state.selectedKey}`} >

                <PivotItem headerText='Pivot 0' itemKey='0' >
                  <GroupTypeChoice
                    departmentSiteTypeName={this.props.departmentSiteTypeName}
                    onChange={this.handleGroupTypeChoice}
                  />
                </PivotItem>

                <PivotItem headerText='Pivot 1' itemKey='1'>
                  {
                    <Activity
                      cancel={this.cancel}
                      context={this.props.context}
                      updateList={this._updateList}
                      hideParentDepartment={this.props.hideParentDepartment}
                    />
                  }
                </PivotItem>

                <PivotItem headerText='Pivot 2' itemKey='2'>
                  {
                    <Department
                      cancel={this.cancel}
                      context={this.props.context}
                      updateList={this._updateList}
                      departmentSiteTypeName={this.props.departmentSiteTypeName}
                      hideParentDepartment={this.props.hideParentDepartment}
                    />
                  }
                </PivotItem>

                <PivotItem headerText='Pivot 3' itemKey='3'>
                  {
                    <Project
                      cancel={this.cancel}
                      context={this.props.context}
                      updateList={this._updateList}
                      hideParentDepartment={this.props.hideParentDepartment}
                    />
                  }
                </PivotItem>

              </Pivot>
            </div>
          </div>
        </Panel>

        <Dialog
          hidden={this.state.hideDialog}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: strings.GroupOrderedSuccesfullyDialog,
            subText: strings.GroupOrderedSuccesfullyDialogSub
          }}
          modalProps={{
            isBlocking: true,
            containerClassName: 'ms-dialogMainOverride'
          }}
        >
          <DialogFooter>
            <DefaultButton onClick={() => { this._closePanel(); this._closeDialog(); }} text={strings.Finish} />
          </DialogFooter>
        </Dialog>

      </div>
    );
  }

  private _showPanel = (): void => {
    this.setState({ showPanel: true });
  }

  private _closePanel = (): void => {
    this.setState({ showPanel: false });
    this.cancel();
  }

  private loadDepartmentPivot(): void {
    this.setState({ selectedKey: 1 });
  }

  private loadActivityPivot(): void {
    this.setState({ selectedKey: 2 });
  }

  private loadProjectPivot(): void {
    this.setState({ selectedKey: 3 });
  }

  private cancel = (): void => {
    this.setState({ selectedKey: 0 });
  }

  private ensureRequestDigest = (): Promise<string> => {
    let { serviceScope, pageContext } = this.props.context;
    // pub.selectedLibrary = {
    //   id: decodedQueryString("SPListId"),
    //   listId: pub.cleanGuid(decodedQueryString("SPListId")),
    //   siteUrl: decodedQueryString("SPSiteUrl"),
    //   urlDir: decodedQueryString("SPListUrlDir"),
    //   url: decodedQueryString("SPSiteUrl") + "/" + decodedQueryString("SPListUrlDir"),
    //   contentTypes: {}
    // };
    // monkey patch #__REQUESTDIGEST element
    var __REQUESTDIGEST = document.getElementById("__REQUESTDIGEST");
    if (!__REQUESTDIGEST) {
      __REQUESTDIGEST = document.createElement("input");
      __REQUESTDIGEST.setAttribute("id", "__REQUESTDIGEST");
      __REQUESTDIGEST.setAttribute("name", "__REQUESTDIGEST");
      __REQUESTDIGEST.setAttribute("type", "hidden");
      const digestCache: IDigestCache = serviceScope.consume(
        DigestCache.serviceKey
      );
      return digestCache
        .fetchDigest(pageContext.web.serverRelativeUrl);
      // .then((digest: string): void => {
      //   // use the digest here
      //   __REQUESTDIGEST.setAttribute("value", digest);
      //   document.body.appendChild(__REQUESTDIGEST);

      //   return digest;
      // })
    }
  }

  public _updateList = (data: any): void => {
    var relativeSiteUrl = this.props.context.pageContext.web.serverRelativeUrl;
    var listName = "Bestillinger";
    var restEndPointUrl = relativeSiteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items";
    var itemType = this._getItemTypeForListName(listName);
    console.log(restEndPointUrl);
    console.log(itemType);
    data['__metadata'] = { 'type': itemType };
    this.ensureRequestDigest().then(requestDigest => {
      $.ajax({
        contentType: 'application/json',
        url: restEndPointUrl,
        type: "POST",
        data: JSON.stringify(data),
        headers: {
          "Accept": "application/json;odata=verbose",
          "content-type": "application/json;odata=verbose",
          "X-RequestDigest": requestDigest
        },
        success: (_data) => {
          this.setState({ updateListCallResult: "success" });
          this._showDialog();
        },
        error: (err) => {
          this.setState({ updateListCallResult: "error" });
          alert(strings.SomethingWentWrong);
          console.log(err);
        }
      });
    });
  }

  private _getItemTypeForListName = (name) => {
    return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
  }

  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

}
