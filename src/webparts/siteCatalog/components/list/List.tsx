import * as React from 'react';
import styles from '../SiteCatalog.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { IOffice365Group } from '../interfaces/office365Group';
import { IContextFilter } from '../interfaces/iContextFilter';
import { Link } from "office-ui-fabric-react/lib/Link";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  CheckboxVisibility,
  IColumn,
  ColumnActionsMode,
  IDetailsRowProps,
  DetailsRow,
  IDetailsList
} from "office-ui-fabric-react/lib/DetailsList";
import { find } from 'core-js/library/fn/array';
import { deepEqual } from "../helpers/deepEqual";
import { T } from '../helpers/translate';
import * as strings from 'SiteCatalogWebPartStrings';
import { WebPartContext } from "@microsoft/sp-webpart-base";
//import { FieldUserRenderer } from '@pnp/spfx-controls-react/lib/FieldUserRenderer';
import { IPrincipal, IUserProfileProperties, IODataKeyValuePair } from '@pnp/spfx-controls-react/lib/common/SPEntities';
import UserProfile from '../user/UserProfile';
import { FieldUserRenderer } from '../user/FieldUserRenderer';
import { FieldUserContainer } from '../user/FieldUserContainer';
import {
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { ListItemContext } from './ListItemContext';
import { SPHttpClient, HttpClientResponse, IGraphHttpClientOptions } from '@microsoft/sp-http';
import { IFavoriteSites } from '../interfaces/IFavoriteSites';

export interface IListState {
  showContext?: boolean;
  contextTarget?: HTMLElement;
  contextFilter?: Array<IContextFilter>;
  activeContextFilters?: any;
  filteredFields?: Array<string>;
  items?: any;
  currentItemId?: string;
  distinctValues?: any;
  sortedColumn?: {
    fieldName?: string;
    direction?: 'asc' | 'desc';
  };
}

export interface IListProps {
    listData: IOffice365Group[];
    filterHandler: (filterData: Array<any>) => void;
    context: WebPartContext;
    emptyResult: boolean;
    sortHandler: (sortedColumn: any) => void;
    favoriteSites: IFavoriteSites;
    siteTypes: string;
    hideParentDepartment: boolean;
}

export default class List extends React.Component<IListProps, IListState> {

  constructor(props: IListProps) {
    super(props);

    this.state = {
      showContext: false,
      contextFilter: [],
      activeContextFilters: [],
      items: [],
      distinctValues: null,
      sortedColumn: { fieldName: 'Title', direction: 'asc' }
    };
  }

  public componentWillReceiveProps(nextProps: IListProps){
    if (
      nextProps.listData &&
      !deepEqual(nextProps.listData, this.state.items)
    ){
      this.insertDistinctValues(nextProps.listData);
    } else if (!nextProps.listData) {
      this.setState({ items: [] });
    }
  }

  public render(): React.ReactElement<IListProps> {
    let { contextTarget, activeContextFilters, showContext, currentItemId, items } = this.state;
    let selectedItem = items.length && find(items, item => item.ID === currentItemId);

    return (
        <div>
            <DetailsList 
              items={this.filterItems()}
              checkboxVisibility={CheckboxVisibility.hidden}
              onItemInvoked={(item, i, ev) => ev && (ev.preventDefault(), ev.stopPropagation())}
              onRenderRow={props => {props.key = props.item.Title + props.item.itemIndex; return <DetailsRow {...props} />;}}
              layoutMode={DetailsListLayoutMode.justified}
              onRenderItemColumn={this.renderListItemColumn}
              columns={this.buildDetailsListColumns()}
              onColumnHeaderContextMenu={this.handleColumnHeaderContextMenu}
            />
            <ListItemContext
              isContextMenuVisible={showContext || false}
              contextTarget={contextTarget as HTMLElement}
              contextFilter={this.state.contextFilter}
              activeContextFilters={activeContextFilters}
              selectedItem={selectedItem}
              handleDismiss={() => {
                this.setState({ showContext: false });
              }}
              handleItemAction={this.handleContextItemAction}
              handleFiltered={(filteredData, activeFilteredData) => {
                this.setState({ contextFilter: filteredData, activeContextFilters: activeFilteredData });
                this.props.filterHandler(activeFilteredData);
              }}
              sortHandler={this.props.sortHandler}
              setSortedColumn={this.setSortedColumn}
              sortedColumn={this.state.sortedColumn}
            />
            {/*this.state.items.length === 0 && !this.props.emptyResult ? <Spinner size={SpinnerSize.large} label={strings.LoadingGroups} ariaLive='assertive' /> : ''*/}
            {this.props.emptyResult ? strings.NoGroupsFound : ''}
        </div>
    );
  }

  public setSortedColumn = (sortedColumn) => {
    this.setState({
      sortedColumn: sortedColumn
    });
  }

  public isSiteAFavorite = (url: string) => {
    if (this.props.favoriteSites) {
      for (let x = 0; x < this.props.favoriteSites.Items.length; x++) {
        if (this.props.favoriteSites.Items[x].Url === url) {
          return true;
        }
      }
      return false;
    }
    return false;
  }

  public getSPURL = (itemId: string) => {
    let url: string;

    let currentItem = this.state.items.filter(item => item.id === itemId);
    let mail = currentItem[0].mail;
    let splittedMail = mail.split('@');
    const siteName = splittedMail[0];

    url = currentItem[0].hasOwnProperty('siteUrl') ? currentItem[0].siteUrl : window.location.protocol + "//" + window.location.host + '/sites/' + siteName;

    return url;
  }

  public getEmailURL = (itemId: string) => {
    let url: string;
    let currentItem = this.state.items.filter(item => item.id === itemId);
    let mail = currentItem[0].mail;

    url = 'https://outlook.office.com/owa/?path=/group/' + mail + '/mail';

    return url;
  }

  public getCalendarUrl = (itemId: string) => {
    let url: string;
    let currentItem = this.state.items.filter(item => item.id === itemId);
    let mail = currentItem[0].mail;

    url = 'https://outlook.office.com/owa/?path=/group/' + mail + '/calendar';

    return url;
  }

  public getPlannerUrl = (itemId: string) => {
    let url:string = this.props.context.pageContext.site.absoluteUrl + '/_layouts/15/groupstatus.aspx?id=' + itemId + '&target=planner';
    return url;
  }

  public handleContextItemAction = actionType => {
    //let url = this.getExternalFormUrl(this.state.currentItemId, actionType);
    let url;

    switch (actionType) {
      case 'sharepoint':
        url = this.getSPURL(this.state.currentItemId);
        break;
      case 'email':
        url = this.getEmailURL(this.state.currentItemId);
        break;
      case 'planner':
        url = this.getPlannerUrl(this.state.currentItemId);
        break;
      case 'calendar':
        url = this.getCalendarUrl(this.state.currentItemId);
        break;
      default:
        break;
    }

    // Open in new window
    let win = window.open(url, '_blank');
    win.focus();
  }

  public handleColumnHeaderContextMenu = (column, ev) => {
    debugger;
    if (!this.state.contextFilter)
      return (console.log(this.state.contextFilter, "no 'contextFilter' is provided"));

    if (!column) return console.assert(column, "no 'column' is provided");
    if (!ev) return console.assert(ev, "no 'ev' is provided");

    if (
      this.state.contextFilter.filter(filter => filter.fieldName === column.key).length === 0
    ) {
      this.getListFilterData(column.key, ev.target as HTMLElement);
    }

    /*this.setState({
      showContext: true,
      contextTarget: ev.target as HTMLElement
    });*/

    return ev && (ev.preventDefault() && ev.stopPropagation());
  }

  public filterItems = () => {

    return this.state.items.map(item => {

      if (!this.props.hideParentDepartment) {
        return {
          Title: item[strings.SchemaName].LabelString01,
          '...': '',
          SiteType: item[strings.SchemaName].ValueString00,
          Manager: item[strings.SchemaName].ValueString03,
          Department: item[strings.SchemaName].ValueString04,
          Mail: item.mail
        };
      } else {
        return {
          Title: item[strings.SchemaName].LabelString01,
          '...': '',
          SiteType: item[strings.SchemaName].ValueString00,
          Manager: item[strings.SchemaName].ValueString03,
          Mail: item.mail
        };
      }

      
      
    });
  }

  public insertDistinctValues = (items) => {
    this.setState({
      items: items
    });
  }

  public itemByIndex = (i: number) => this.state.items[i];

  public renderListItemColumn = (item:IOffice365Group, i, col: IColumn) => {
    //console.log(col);
    const currentItem = this.itemByIndex(i);

    if (col.fieldName === 'Title') {
      //col.isSorted = true;
      const mail = currentItem.mail;
      const splittedMail = mail.split('@');
      let url = currentItem.hasOwnProperty('siteUrl') ? currentItem.siteUrl : window.location.protocol + "//" + window.location.host + '/sites/' + splittedMail[0];

      let privateGroup = <i className={styles.groupIcons + ' ms-Icon ms-Icon--Lock'} aria-hidden="true"></i>;
      let favoritedGroup = <i className={styles.groupIcons + ' ms-Icon ms-Icon--FavoriteStar'} aria-hidden="true"></i>;
      const isFavorite = this.isSiteAFavorite(url);

      return (
        <div>
          <Link onClick={evt => {
              evt.preventDefault();
              let win = window.open(url, '_blank');
              win.focus();
            }
          }>
            {currentItem[strings.SchemaName].LabelString01}
          </Link>
          {currentItem.visibility !== 'Public' ? privateGroup : ''}
          {isFavorite ? favoritedGroup : ''}
        </div>
      );

    } else if(col.key === '...'){
      col.className = "rowSelector " + styles.alignCenter;
      col.isIconOnly = true;
      col.iconClassName = "NoIcon";
      col.name = "";
      col.columnActionsMode = ColumnActionsMode.disabled;

      let contextFilter = [
        {
          fieldName: "more",
          filterText: "",
          filterValue: "",
          isFiltered: false
        }
      ];
      return (
        <div
          onClick={ev => {
            this.setState({
              showContext: true,
              contextTarget: ev.target as HTMLElement,
              contextFilter: contextFilter,
              currentItemId: currentItem.id
            });
          }}
          className={styles.moreButton}
        >
          ...
        </div>
      );
    } else if (col.key === 'SiteType') {
      //console.log('HELLO');
      //console.log(item[strings.SchemaName].ValueString00);
      return currentItem[strings.SchemaName].ValueString00;
    } else if (col.key === 'Manager'){
      //return currentItem.userProfile.DisplayName;
      return (<FieldUserContainer key={currentItem.id} group={this.state.items[i]} />);
      //return (<UserProfile context={this.props.context} userLoginName={currentItem[strings.SchemaName].ValueString03} />);

      //return currentItem[strings.SchemaName].ValueString03;
    } else if (col.key === 'Mail' ) {
      return (<a href={'mailto:'+currentItem.mail}>{currentItem.mail}</a>);
    } else if (col.key === 'Department' && !this.props.hideParentDepartment) {
      return currentItem[strings.SchemaName].ValueString04;
    }
  }

   /**
   * builds IColumn[] for DetailsList component.
   * @returns {Array<IColumn>}
   *
   */
  public buildDetailsListColumns = ():IColumn[] => {

    if (this.filterItems().length === 0) {
      if (!this.props.hideParentDepartment) {
        return [
          this.colNameToIColumn("Title"),
          this.colNameToIColumn("..."),
          this.colNameToIColumn("SiteType"),
          this.colNameToIColumn("Manager"),
          this.colNameToIColumn("Department"),
          this.colNameToIColumn("Mail"),
        ] as IColumn[];
      } else {
        return [
          this.colNameToIColumn("Title"),
          this.colNameToIColumn("..."),
          this.colNameToIColumn("SiteType"),
          this.colNameToIColumn("Manager"),
          this.colNameToIColumn("Mail"),
        ] as IColumn[];
      }
    }
      

    let columns: IColumn[] = Object.keys(this.filterItems()[0]).map(
      this.colNameToIColumn
    );

    return columns;
  }

  /**
   * TODO: Implement filtering onclick
   */
  public colNameToIColumn = (colName: string) => {
    let colProps = {
      key: colName,
      fieldName: colName,
      name: T(colName),
      isResizeable: true,
      minWidth: 60
    };

    if (colName === "...") {
      (colProps as IColumn).maxWidth = 30;
      (colProps as IColumn).calculatedWidth = 30;
      (colProps as IColumn).minWidth = 30;
      (colProps as IColumn).isResizable = false;
      (colProps as IColumn).isCollapsable = false;
      (colProps as IColumn).isIconOnly = true;
    } else if (colName === "Department" && !this.props.hideParentDepartment) {
      (colProps as IColumn).minWidth = 100;
      (colProps as IColumn).maxWidth = 180;
    } else if (colName === "Manager") {
      (colProps as IColumn).minWidth = 120;
      (colProps as IColumn).maxWidth = 220;
    } else if (colName === "SiteType") {
      (colProps as IColumn).minWidth = 100;
      (colProps as IColumn).maxWidth = 150;
    } else if (colName === "Mail") {
      (colProps as IColumn).minWidth = 140;
      (colProps as IColumn).maxWidth = 150;
    } else if (colName === "Title") {
      (colProps as IColumn).minWidth = 140;
      (colProps as IColumn).maxWidth = 230;
    }

    let filterFieldName =
      this.state.activeContextFilters &&
      this.state.activeContextFilters.filter(
        f => f.fieldName === colName
      );

    (colProps as IColumn).isFiltered =
      filterFieldName && filterFieldName.length ? true : false;

    if(this.state.sortedColumn && this.state.sortedColumn.fieldName === colName) {
      if (this.state.sortedColumn.direction === 'asc') {
        (colProps as IColumn).iconName = "SortUp";
      } else if (this.state.sortedColumn.direction === 'desc') {
        (colProps as IColumn).iconName = "Down";
      }
    }

    if (
      [
        "Title",
        "Mail",
        "Manager"
      ].indexOf(colName) >= 0
    ) {

      

      (colProps as IColumn).onColumnClick = (ev, column) => {
        let direction;

        if (!this.state.sortedColumn || this.state.sortedColumn.fieldName !== column.key) {
          direction = 'asc';
        } else if (this.state.sortedColumn.direction === 'asc' && this.state.sortedColumn.fieldName === column.key) {
          direction = 'desc';
        } else if (this.state.sortedColumn.direction === 'desc' && this.state.sortedColumn.fieldName === column.key) {
          direction = 'asc';
        }

        let sortedColumn = {
          fieldName: column.key,
          direction: direction
        };

        this.setState({
          sortedColumn: sortedColumn
        });

        this.props.sortHandler(sortedColumn);

        return ev && (ev.preventDefault(), ev.stopPropagation());
      };
    }

    // Enable filtering
    if (
      [
        "SiteType",
        "Department"
      ].indexOf(colName) >= 0
    ) {
      (colProps as IColumn).iconName = "ChevronDown";
      (colProps as IColumn).iconClassName = styles["floatRight"];
      (colProps as IColumn).onColumnClick = (ev, column) => {
        if (!this.state.contextFilter)
          return console.log(
            this.state.contextFilter,
            "no 'contextFilter' is provided"
          );
        if (!column) return console.log(column, "no 'column' is provided");
        if (!ev) return console.log(ev, "no 'event' is provided");

        if (
          this.state.contextFilter.filter(
            filter => filter.fieldName === column.key
          ).length === 0
        ) {
          this.getListFilterData(column.key, ev.target as HTMLElement);
          // this.setState({
          //   showContext: true,
          //   contextTarget: ev.target as HTMLElement,
          //   contextFilter: []
          // });
        } else {
          this.setState({
            showContext: true,
            contextTarget: ev.target as HTMLElement
          });
        }

        return ev && (ev.preventDefault(), ev.stopPropagation());
      };
    }

    return colProps as IColumn;

  }

  private _sortItems = (items: any, sortBy: string, subProp:string, descending = false): any[] => {
    if (subProp === null) {
      if (descending) {
        return items.sort((a: any, b: any) => {
          if (a[sortBy].toLowerCase() < b[sortBy].toLowerCase()) {
            return 1;
          }
          if (a[sortBy].toLowerCase() > b[sortBy].toLowerCase()) {
            return -1;
          }
          return 0;
        });
      } else {
        return items.sort((a: any, b: any) => {
          if (a[sortBy].toLowerCase() < b[sortBy].toLowerCase()) {
            return -1;
          }
          if (a[sortBy].toLowerCase() > b[sortBy].toLowerCase()) {
            return 1;
          }
          return 0;
        });
      }
    } else {
      if (descending) {
        return items.sort((a: any, b: any) => {
          if (a[sortBy][subProp].toLowerCase() < b[sortBy][subProp].toLowerCase()) {
            return 1;
          }
          if (a[sortBy][subProp].toLowerCase() > b[sortBy][subProp].toLowerCase()) {
            return -1;
          }
          return 0;
        });
      } else {
        return items.sort((a: any, b: any) => {
          if (a[sortBy][subProp].toLowerCase() < b[sortBy][subProp].toLowerCase()) {
            return -1;
          }
          if (a[sortBy][subProp].toLowerCase() > b[sortBy][subProp].toLowerCase()) {
            return 1;
          }
          return 0;
        });
      }
    }
  }

  public isFiltered = (fieldName: string, filterText: string) => {
    let found = false;

    this.state.activeContextFilters.forEach(item => {
      item.fieldName === fieldName && item.filterText == filterText ? found = true : found = false;
    });

    return found;
  }

  public setContextFilter = (values:any, columnName:string, evTarget?: HTMLElement) => {
    let filter:IContextFilter[] = [];

    for (let x = 0; x < values.length; x++) {
      let filtered = this.isFiltered(columnName, values[x]);

      let currentFilter:IContextFilter = {
        fieldName: columnName,
        filterValue: values[x],
        filterText: values[x],
        isFiltered: filtered
      };
      filter.push(currentFilter);
    }

    if (evTarget) {
      this.setState({
          showContext: true,
          contextTarget: evTarget as HTMLElement,
          contextFilter: filter
        });
    } else {
      this.setState({
        contextFilter: filter
      });
    }
    
  }

  // Should get filter data and update this.state.contextFilter with array of IContextFilter's
  public getListFilterData = (columnName: string, evTarget?: HTMLElement) => {
    //var dfd = $.Deferred();

    // const val = (filter) => `${typeof filter.filterValue === 'string' ? '\'' : ''}${filter.filterValue}${typeof filter.filterValue === 'string' ? '\'' : ''}`;
    let listName = "Bestillinger";
    let choiceFieldName = "KDTOParentDepartment";
    let choiceFieldId = "5639ffa9-c62e-4513-b7c7-ccca2b5e92c2";

    let handler = this;

    if (columnName === 'Department') {
      // get departments from list
      this.props.context.spHttpClient.get(`${this.props.context.pageContext.site.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/fields('${choiceFieldId}')/Choices`, SPHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        return response.json();
      }).then((result: any) => {

        if (result.value) {
          console.log(result.value);
          handler.setContextFilter(result.value, columnName, evTarget);
        }
  
      });

    } else if (columnName === 'SiteType') {
      // get available content types from list
      /*this.props.context.spHttpClient.get(`${this.props.context.pageContext.site.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/contenttypes`, SPHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        return response.json();
      }).then((result: any) => {

        if (result.value) {
          handler.setContextFilter(result.value.map(obj => {
            return obj.Name;
          }), columnName);
        }
  
      });*/
      let siteTypes = this.props.siteTypes.split(',');
      handler.setContextFilter(siteTypes, columnName, evTarget);
    }

    // Update filterState


    /*let fieldInternalName;

    if (columnName === 'Department') {
      fieldInternalName = 'KDTO.'
    }
    var filterColElement = this;
    var filterColumnResource =
      "RenderListFilterData?ExcludeFieldFilteringHtml=true&FieldInternalName='" +
      columnName +
      "'&FilterName=ID&FilterMultiValue=";
    var url = `${this.props.context.pageContext.site.absoluteUrl}/_api/web/lists/getByTitle('Tiltak')/${filterColumnResource}`;
    

    url += "*";
    this.requestFilterData(siteInfo.DSVUrl, url, columnName).then((contextFilterPartition: Array<IContextFilter>) => {
      return this.updateFilterState(dfd, contextFilterPartition);
    });*/
    

  }



}
