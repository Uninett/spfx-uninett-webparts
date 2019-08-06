import * as React from 'react';
import * as strings from 'SiteCatalogWebPartStrings';
import styles from '../SiteCatalog.module.scss';
import { ISiteCatalogProps } from './ISiteCatalogProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IOffice365Group } from '../interfaces/office365Group';
import { AadHttpClientFactory, AadHttpClient, GraphHttpClient, HttpClientResponse, IGraphHttpClientOptions } from '@microsoft/sp-http';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import List from '../list/List';
import UserProfile from '../user/UserProfile';
import { FieldUserRenderer } from '@pnp/spfx-controls-react/lib/FieldUserRenderer';
import { IPrincipal, IUserProfileProperties, IODataKeyValuePair } from '@pnp/spfx-controls-react/lib/common/SPEntities';
import { UserProfileService } from '../user/UserProfileService';
import { IUserProfile } from '../interfaces/IUserProfile';
import { SearchBoxContainer } from '../searchBox/SearchBoxContainer';
import { IContextFilter } from '../interfaces/iContextFilter';
import { IFavoriteSites } from '../interfaces/IFavoriteSites';
import { NavigationBar } from '../navigationBar/NavigationBar';
import { ToolBar } from '../toolbar/Toolbar';
require('../scss/custom.module.scss');
import { getUserProfileProperty } from '../helpers/userProfile';
import { FavoriteSitesService } from '../helpers/FavoriteSitesService';
import { IGroupSiteUrl } from '../interfaces/IGroupSiteUrl';
import { GroupSiteUrlService } from '../../services/GroupSiteUrlService';

export interface ISiteCatalogState {
  isLoading: boolean;
  listData?: Array<IOffice365Group[]>;
  unPartitionedListData?: IOffice365Group[];
  filter?: Array<{
    fieldName?: string;
    direction?: 'asc' | 'desc';
    filterValue?: any;
  }>;
  emptyResult: boolean;
  sortedColumn?: {
    fieldName?: string;
    direction?: 'asc' | 'desc';
  };
  navPages: Array<INavPage>;
  currentPage: number;
  memberGroups: Array<any>;
  searchString: string;
  favoriteSites: IFavoriteSites;
}

export default class SiteCatalog extends React.Component<ISiteCatalogProps, ISiteCatalogState> {

  constructor(props) {
    super(props);

    this.state = {
      filter: [],
      emptyResult: false,
      navPages: [],
      currentPage: 0,
      memberGroups: [],
      listData: [],
      unPartitionedListData: [],
      sortedColumn: { fieldName: 'Title', direction: 'asc' },
      isLoading: false,
      searchString: '',
      favoriteSites: null
    };
  }

  public componentWillMount() {
    // Fetch data from graph

    this.getMemberGroupsAndFetchData();
    this.getFavoriteSites();
  }

  public render(): React.ReactElement<ISiteCatalogProps> {

    return (
      <div className={styles.siteCatalog}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={'ms-Grid-col ms-sm6 ms-md7 ms-lg7'}>
              <h1 className={'ms-font-xxl'}>{strings.WebPartTitle}</h1>
            </div>
            <div className={'ms-Grid-col ms-sm6 ms-md5 ms-lg5'}>
              <SearchBoxContainer searchHandler={this.searchHandler} clearHandler={this.searchClearHandler} />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <ToolBar
                currPage={(this.state.currentPage + 1).toString()}
                totalPages={this.state.listData.length > 0 ? this.state.listData.length.toString() : '1'}
                loading={this.state.isLoading}
                handleNewClick={this.handleNewClick}
                enableNext={this.state.listData.length > 0 && this.state.currentPage < this.state.listData.length - 1 ? true : false}
                enablePrev={this.state.listData.length > 0 && this.state.currentPage !== 0 ? true : false}
                navigateHandler={this.navigateHandler}
                showNewButton={this.props.showNewButton}
              />
              <List
                sortHandler={this.sortHandler}
                listData={this.state.listData[this.state.currentPage]}
                filterHandler={this.refreshListOnFilter}
                context={this.props.context}
                emptyResult={this.state.emptyResult}
                favoriteSites={this.state.favoriteSites}
                siteTypes={this.props.siteTypes}
                hideParentDepartment={this.props.hideParentDepartment}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }

  public getFavoriteSites = () => {
    let handler = this;
    let favoriteSitesService: FavoriteSitesService = new FavoriteSitesService(this.props.context);
    favoriteSitesService.getFavoriteSites().then((result) => {
      if (result.hasOwnProperty('Items')) {
        handler.setState({
          favoriteSites: result
        });
      } else {
        console.log('Could not fetch favorite sites.');
      }

    }).catch((error) => {
      console.log('Could not fetch favorite sites.');
    });
  }

  // Fetches only groups that the user is a member of, to deal with permissions. Then we fetch all groups to get metadata.
  public getMemberGroupsAndFetchData = ():void => {

    let handler = this;

    this.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient): void => {
        client
          .get("https://graph.microsoft.com/v1.0/users/" + this.props.context.pageContext.user.loginName + "/memberOf/$/microsoft.graph.group?$filter=groupTypes/any(a:a eq 'unified')", AadHttpClient.configurations.v1)
          .then((response: HttpClientResponse) => {
            if (response.ok) {
              return response.json();
            } else {
              console.warn(response.statusText);
            }
          })
          .then((result: any) => {
            if (result && result.hasOwnProperty('value')) {
              handler.setState({ memberGroups: result.value });
      
              // Fetch all groups
              this.loadSiteList();
            }
          });
    });

/*
    // Fetch member groups
    this.props.context.graphHttpClient.get("v1.0/users/" + this.props.context.pageContext.user.loginName + "/memberOf/$/microsoft.graph.group?$filter=groupTypes/any(a:a eq 'unified')", GraphHttpClient.configurations.v1).then((response: HttpClientResponse) => {
      if (response.ok) {
        return response.json();
      } else {
        console.warn(response.statusText);
      }
    }).then((result: any) => {
      if (result && result.hasOwnProperty('value')) {
        handler.setState({ memberGroups: result.value });

        // Fetch all groups
        this.loadSiteList();
      }
    });
*/

  }

  // Controls which result page to show in list
  public navigateHandler = (ev, direction) => {
    if (direction === 'next') {
      this.setState({
        isLoading: true
      });
      let currentPage = this.state.currentPage + 1;
      this.fetchMissingUserprofiles(this.state.listData, this.state.unPartitionedListData, null, currentPage);
      /**/
    } else if (direction === 'back') {
      this.setState({
        isLoading: true
      });
      let currentPage = this.state.currentPage - 1;
      this.fetchMissingUserprofiles(this.state.listData, this.state.unPartitionedListData, null, currentPage);

    }
  }

  private fetchMissingUserprofiles = (partitionedListData: Array<IOffice365Group[]>, unPartitionedListData: IOffice365Group[], sortedColumn?, nextPage?: number) => {

    let promises: Promise<IUserProfile>[] = [];
    let userProfileService: UserProfileService;

    let distinctValues: string[] = [];
    let handler = this;
    let sortedCol = sortedColumn ? sortedColumn : this.state.sortedColumn;


    let currentPage = nextPage !== null ? nextPage : this.state.currentPage;

    // Fetching users
    for (let y = 0; y < partitionedListData[currentPage].length; y++) {
      let userLoginName = partitionedListData[currentPage][y][strings.SchemaName].ValueString03;

      if (!partitionedListData[currentPage][y].hasOwnProperty('userProfile') || !partitionedListData[currentPage][y]['userProfile']) {
        distinctValues.push(userLoginName);

        userProfileService = new UserProfileService(this.props.context, userLoginName);
        promises.push(userProfileService.getUserProfileProperties());
      }
    }


    Promise.all(promises).then((values) => {

      if (values.length === 0) {
        this.setState({
          listData: partitionedListData,
          unPartitionedListData: unPartitionedListData,
          sortedColumn: sortedCol,
          isLoading: false,
          emptyResult: false,
          currentPage: currentPage
        });
      } else {
        for (let x = 0; x < partitionedListData[currentPage].length; x++) {

          if ((partitionedListData[currentPage][x].hasOwnProperty('userProfile') || partitionedListData[currentPage][x]['userProfile'])) {
            continue;
          }

          let userLoginName = partitionedListData[currentPage][x][strings.SchemaName].ValueString03;
          let userData;

          for (let z = 0; z < values.length; z++) {

            if (!values[z].hasOwnProperty('UserProfileProperties') || !values[z].UserProfileProperties) {
              continue;
            }

            if (getUserProfileProperty(values[z].UserProfileProperties, 'UserName') === userLoginName) {
              for (let y = 0; y < values[z].UserProfileProperties.length; y++) {
                if (values[z].UserProfileProperties[y].Key === 'PictureURL') {
                  let imageUrl = this.props.context.pageContext.site.absoluteUrl + '/_layouts/15/userphoto.aspx?size=M&username=' + userLoginName;
                  values[z].UserProfileProperties[y].Value = imageUrl;
                  break;
                }
              }
              userData = values[z];
              break;
            }
          }

          partitionedListData[currentPage][x].userProfile = userData;
        }

        // Merge partitions to generate complete list data, used for sorting later on
        let newUnPartitionatedListData: IOffice365Group[] = [];
        for (let y = 0; y < partitionedListData.length; y++) {
          if (y === 0) {
            newUnPartitionatedListData = partitionedListData[0];
          } else {
            newUnPartitionatedListData = newUnPartitionatedListData.concat(partitionedListData[y]);
          }
        }



        this.setState({
          listData: partitionedListData,
          unPartitionedListData: newUnPartitionatedListData,
          sortedColumn: sortedCol,
          isLoading: false,
          emptyResult: false,
          currentPage: currentPage
        }, () => this.getSiteUrlsAsync(partitionedListData, newUnPartitionatedListData, currentPage));
      }

    });
  }

  public sortHandler = (sortedColumn) => {

    if (!sortedColumn || this.state.emptyResult || this.state.listData.length === 0) {
      this.setState(
        {
          sortedColumn: {
            direction: 'asc',
            fieldName: 'Name'
          }
        }
      );
      return;
    }

    let direction;

    let descending;
    let sortBy;
    let subProp = null;

    if (sortedColumn.direction === 'asc') {
      descending = false;
    } else {
      descending = true;
    }

    if (sortedColumn.fieldName === 'Title') {
      sortBy = 'extvcs569it_InmetaGenericSchema';
      subProp = 'LabelString01';
    } else if (sortedColumn.fieldName === 'Mail') {
      sortBy = 'mail';
    } else if (sortedColumn.fieldName === 'Manager') {
      sortBy = 'extvcs569it_InmetaGenericSchema';
      subProp = 'LabelString03';
    } else if (sortedColumn.fieldName === 'SiteType') {
      sortBy = 'extvcs569it_InmetaGenericSchema';
      subProp = 'ValueString00';
    } else if (sortedColumn.fieldName === 'Department') {
      sortBy = 'extvcs569it_InmetaGenericSchema';
      subProp = 'ValueString04';
    }

    let _sortedColumn = {
      direction: sortedColumn.direction,
      fieldName: sortedColumn.fieldName
    };
    let sortedItems = this._sortItems(this.state.unPartitionedListData, sortBy, subProp, descending);
    let partitionatedSortedData = this.partitionateListData(sortedItems);

    this.setState({
      isLoading: true
    });

    this.fetchMissingUserprofiles(partitionatedSortedData, sortedItems, _sortedColumn, this.state.currentPage);
  }

  private sortInternal = (sortedColumn, listData) => {

    if (!sortedColumn) {
      return listData;
    }
    let direction;

    let descending;
    let sortBy;
    let subProp = null;

    if (sortedColumn.direction === 'asc') {
      descending = false;
    } else {
      descending = true;
    }

    if (sortedColumn.fieldName === 'Title') {
      sortBy = 'extvcs569it_InmetaGenericSchema';
      subProp = 'LabelString01';
    } else if (sortedColumn.fieldName === 'Mail') {
      sortBy = 'mail';
    } else if (sortedColumn.fieldName === 'Manager') {
      sortBy = 'extvcs569it_InmetaGenericSchema';
      subProp = 'LabelString03';
    } else if (sortedColumn.fieldName === 'SiteType') {
      sortBy = 'extvcs569it_InmetaGenericSchema';
      subProp = 'ValueString00';
    } else if (sortedColumn.fieldName === 'Department') {
      sortBy = 'extvcs569it_InmetaGenericSchema';
      subProp = 'ValueString04';
    }

    let sortedItems = this._sortItems(listData, sortBy, subProp, descending);

    return sortedItems;
  }

  private _sortItems = (items: any, sortBy: string, subProp: string, descending = false): any[] => {
    
    if (subProp === null) {
      if (descending) {
        return items.sort((a: any, b: any) => {
          if (!a.hasOwnProperty(sortBy)) {
            return 1;
          }
          if (!b.hasOwnProperty(sortBy)) {
            return -1;
          }

          if (!a[sortBy]) {
            return 1;
          }

          if (!b[sortBy]) {
            return -1;
          }

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
          if (!a.hasOwnProperty(sortBy)) {
            return 1;
          }
          if (!b.hasOwnProperty(sortBy)) {
            return -1;
          }

          if (!a[sortBy]) {
            return 1;
          }

          if (!b[sortBy]) {
            return -1;
          }

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


          if (!a.hasOwnProperty(sortBy)) {
            return 1;
          }
          if (!b.hasOwnProperty(sortBy)) {
            return -1;
          }

          if (!a[sortBy]) {
            return 1;
          }

          if (!a[sortBy].hasOwnProperty(subProp)) {
            return 1;
          }
          
          if (!a[sortBy][subProp]) {
            return 1;
          }

          if (!b[sortBy]) {
            return -1;
          }

          if (!b[sortBy].hasOwnProperty(subProp)) {
            return -1;
          }
          
          if (!b[sortBy][subProp]) {
            return -1;
          }

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
          if (!a.hasOwnProperty(sortBy)) {
            return 1;
          }
          if (!b.hasOwnProperty(sortBy)) {
            return -1;
          }

          if (!a[sortBy]) {
            return 1;
          }

          if (!a[sortBy].hasOwnProperty(subProp)) {
            return 1;
          }
          
          if (!a[sortBy][subProp]) {
            return 1;
          }

          if (!b[sortBy]) {
            return -1;
          }

          if (!b[sortBy].hasOwnProperty(subProp)) {
            return -1;
          }
          
          if (!b[sortBy][subProp]) {
            return -1;
          }

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

  public searchHandler = (keyWord: string) => {
    this.setState({
      listData: [],
      unPartitionedListData: [],
      navPages: [],
      currentPage: 0,
      emptyResult: false,
      searchString: keyWord
    }, this.loadSiteList);
  }

  public searchClearHandler = () => {
    this.setState({
      listData: [],
      unPartitionedListData: [],
      navPages: [],
      currentPage: 0,
      emptyResult: false,
      searchString: ''
    }, this.loadSiteList);
  }

  // Groups different filters together. Sitetype should be first
  public groupFilter = (filters: any[]) => {
    
    let temp = [];
    for (let x = 0; x < filters.length; x++) {
      if (filters[x].fieldName === 'SiteType') {
        let currSiteType = filters[x];
        filters.splice(x, 1);
        temp.push(currSiteType);
      }
    }

    const result = temp.concat(filters);
    return result;
  }

  public loadSiteList = (removedParam?: string, linkUrl?: string) => {
    let handler = this;
    let url;
    let newReq;
    let searchString = this.state.searchString;
    let groupedFilters = [];
    
    this.setState({ isLoading: true });

    if (!linkUrl) {

      let filter = "$filter=";
      newReq = true;
      console.log("Using " + this.props.searchType + " search type");
      // Apply filters
      if ((this.state.filter.length === 0 && this.props.searchType === 'javascript') || (this.props.searchType === 'graph' && !searchString && this.state.filter.length === 0)) {
        filter += "extvcs569it_InmetaGenericSchema/KeyString00 eq 'SiteType'";
      } else {

        if (this.props.searchType === 'graph' && searchString){
          filter += `
          (startswith(extvcs569it_InmetaGenericSchema/ValueString00, '`+searchString+`') or 
          startswith(extvcs569it_InmetaGenericSchema/LabelString03, '`+searchString+`') or 
          startsWith(extvcs569it_InmetaGenericSchema/LabelString01, '`+searchString+`') or 
          startsWith(extvcs569it_InmetaGenericSchema/ValueString04, '`+searchString+`'))
          `;
        }

        if (this.state.filter.length > 0) {
          
          // Group array -> department first, then array to deal with 'and' and 'or'
          groupedFilters = this.groupFilter(this.state.filter);

          if (this.props.searchType === 'graph' && searchString) {
            filter += ` and (`;
          }

          for (let x = 0; x < groupedFilters.length; x++) {
            if (x === 0) {
              filter += '(';
            }

            if (groupedFilters[x].fieldName === 'SiteType') {

              // Insert or if previous field is of same type
              if (x !== 0) {
                if (groupedFilters[x - 1].fieldName === 'SiteType') {
                  filter += ' or ';
                }
              }

              filter += `extvcs569it_InmetaGenericSchema/ValueString00 eq '${groupedFilters[x].filterValue}'`;

              if (x !== groupedFilters.length - 1) {
                if (groupedFilters[x + 1].fieldName !== 'SiteType') {
                  filter += ') and (';
                }
              }
            } else if (groupedFilters[x].fieldName === 'Department') {

              if (x !== 0) {
                if (groupedFilters[x - 1].fieldName === 'Department') {
                  filter += ' or ';
                }
              }
              filter += `extvcs569it_InmetaGenericSchema/ValueString04 eq '${groupedFilters[x].filterValue}'`;
            }

            if (x === groupedFilters.length - 1) {
              filter += ')';
            }

          }

          if (this.props.searchType === 'graph' && searchString) {
            filter += `)`;
          }

        }

      }





      let select = '&$select=displayName,id,extvcs569it_InmetaGenericSchema,mail,visibility';
      //let limit = this.props.listRows ? '&$top=' + this.props.listRows : '';
      let limit = '&$top=999';

      // commented out temporarily
      //url = filter + select + limit;
      url = filter + select + limit;

    } else {
      newReq = false;
      url = linkUrl;
    }

    this.props.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient): void => {
        client
          .get("https://graph.microsoft.com/v1.0/groups?" + url, AadHttpClient.configurations.v1)
          .then((response: HttpClientResponse) => {
            if (response.ok) {
              return response.json();
            } else {
              console.warn(response.statusText);
            }
          })
          .then((result: any) => {
            if (result && result.value.length > 0) {
              // Filter result
              //let permissonTrimmedResult: Array<any> = result.value.filter(group => { return this.state.memberGroups.some(memberGroup => { return (group.id === memberGroup.id || group.visibility === 'Public') })});
              let permissonTrimmedResult: Array<any> = [];
              permissonTrimmedResult = this.getPermissionTrimmedResult(result.value);
      
              if (permissonTrimmedResult && permissonTrimmedResult.length > 0) {
                let navPages = [];
      
                handler.fetchUserProfiles(permissonTrimmedResult, navPages, groupedFilters);
              } else {
                handler.setState({
                  emptyResult: true,
                  currentPage: 0,
                  filter: groupedFilters,
                  listData: [],
                  unPartitionedListData: [],
                  navPages: [],
                  isLoading: false
                });
              }
      
      
            } else {
              if (result.error) {
                console.error(result.message);
              } else if (this.state.filter && result.value.length == 0) {
                handler.setState({
                  emptyResult: true,
                  currentPage: 0,
                  filter: groupedFilters,
                  listData: [],
                  unPartitionedListData: [],
                  navPages: [],
                  isLoading: false
                });
              } else {
                handler.setState({
                  emptyResult: true,
                  filter: groupedFilters,
                  currentPage: 0,
                  navPages: [],
                  isLoading: false
                });
              }
      
            }
          });
      });

/*    // Query for all groupos on the tenant using Microsoft Graph.
    this.props.context.graphHttpClient.get('v1.0/groups?' + url, GraphHttpClient.configurations.v1).then((response: HttpClientResponse) => {
      if (response.ok) {
        return response.json();
      } else {
        console.warn(response.statusText);
      }
    }).then((result: any) => {
      if (result && result.value.length > 0) {
        // Filter result
        //let permissonTrimmedResult: Array<any> = result.value.filter(group => { return this.state.memberGroups.some(memberGroup => { return (group.id === memberGroup.id || group.visibility === 'Public') })});
        let permissonTrimmedResult: Array<any> = [];
        permissonTrimmedResult = this.getPermissionTrimmedResult(result.value);

        if (permissonTrimmedResult && permissonTrimmedResult.length > 0) {
          let navPages = [];

          handler.fetchUserProfiles(permissonTrimmedResult, navPages, groupedFilters);
        } else {
          handler.setState({
            emptyResult: true,
            currentPage: 0,
            filter: groupedFilters,
            listData: [],
            unPartitionedListData: [],
            navPages: [],
            isLoading: false
          });
        }


      } else {
        if (result.error) {
          console.error(result.message);
        } else if (this.state.filter && result.value.length == 0) {
          handler.setState({
            emptyResult: true,
            currentPage: 0,
            filter: groupedFilters,
            listData: [],
            unPartitionedListData: [],
            navPages: [],
            isLoading: false
          });
        } else {
          handler.setState({
            emptyResult: true,
            filter: groupedFilters,
            currentPage: 0,
            navPages: [],
            isLoading: false
          });
        }

      }

    });
  */
  }

  public getPermissionTrimmedResult = (result: Array<any>) => {
    let permissionTrimmedRes = [];

    if (result.length === 0) {
      return permissionTrimmedRes;
    }

    // Add public groups
    for (let x = 0; x < result.length; x++) {
      if (result[x].visibility === 'Public') {
        permissionTrimmedRes.push(result[x]);
      } else if (this.groupExist(this.state.memberGroups, result[x].id)) {
        permissionTrimmedRes.push(result[x]);
      }
    }

    return permissionTrimmedRes;
  }

  // Sets navigation links for result
  public setNavLinks = (result, newRequest: boolean, requestUrl?: string, paginationLink?: string) => {

    let pages;

    if (newRequest) {
      pages = [];
    } else {
      pages = [...this.state.navPages];
    }


    let nextLink;
    if (result['@odata.nextLink']) {
      nextLink = result['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0/groups?', '');
    }


    if (pages.length === 0 && nextLink) {
      let page: INavPage = {
        nextLink: nextLink,
        currentLink: requestUrl,
        backLink: ''
      };

      pages.push(page);
    } else if (pages.length === 0 && !nextLink) {
      let page: INavPage = {
        nextLink: '',
        currentLink: requestUrl,
        backLink: ''
      };

      pages.push(page);
    } else if (!pages[this.state.currentPage]) {

      let page: INavPage = {
        nextLink: '',
        currentLink: requestUrl,
        backLink: pages[this.state.currentPage - 1].currentLink
      };

      if (nextLink)
        page.nextLink = nextLink;

      pages.push(page);
    }

    return pages;
  }

  public propertyExist = (arr: string[], val: string): boolean => {
    for (let x = 0; x < arr.length; x++) {
      if (arr[x] === val) return true;
    }

    return false;
  }

  public groupExist = (arr: Array<any>, id: string): boolean => {
    for (let x = 0; x < arr.length; x++) {
      if (arr[x].id === id) return true;
    }

    return false;
  }

  public search = (data: IOffice365Group[]): IOffice365Group[] => {
    let searchResult: IOffice365Group[] = [];
    const searchWord = this.state.searchString.toLowerCase();

    data.forEach((element, index) => {
      let pushed: boolean = false;

      if (element['extvcs569it_InmetaGenericSchema']) {

        if (element['extvcs569it_InmetaGenericSchema']['LabelString01']) {
          if (element['extvcs569it_InmetaGenericSchema']['LabelString01'].toLowerCase().indexOf(searchWord) !== -1) {
            searchResult.push(element);
            pushed = true;
          }
        }
        
        if (element['extvcs569it_InmetaGenericSchema']['LabelString03'] && !pushed) {
          if (element['extvcs569it_InmetaGenericSchema']['LabelString03'].toLowerCase().indexOf(searchWord) !== -1) {
            searchResult.push(element);
            pushed = true;
          }
        }
        
        if (element['extvcs569it_InmetaGenericSchema']['ValueString04'] && !pushed) {
          if (element['extvcs569it_InmetaGenericSchema']['ValueString04'].toLowerCase().indexOf(searchWord) !== -1) {
            searchResult.push(element);
            pushed = true;
          }
        }
        
        if (element['extvcs569it_InmetaGenericSchema']['ValueString00'] && !pushed) {
          if (element['extvcs569it_InmetaGenericSchema']['ValueString00'].toLowerCase().indexOf(searchWord) !== -1) {
            searchResult.push(element);
            pushed = true;
          }
        }
        
        if (element['displayName'].toLowerCase().indexOf(searchWord) !== -1 && !pushed) {
          searchResult.push(element);
          pushed = true;
        }

      } else {
        if (element['displayName'].toLowerCase().indexOf(searchWord) !== -1) {
          searchResult.push(element);
          pushed = true;
        }
      }

    });

    return searchResult;
  }

  public fetchUserProfiles = (data: IOffice365Group[], navPages: INavPage[], groupedFilters: Array<any>) => {

    let listData: IOffice365Group[];
    let promises: Promise<IUserProfile>[] = [];
    let userProfileService: UserProfileService;

    let distinctValues: string[] = [];
    let handler = this;

    let partitionedListData;

    // Handle search
    if (this.props.searchType === 'javascript' && this.state.searchString) {
      listData = this.search(data);
    } else {
      listData = data;
    }

    if (this.state.sortedColumn) {
      let sortedListData = this.sortInternal(this.state.sortedColumn, listData);
      partitionedListData = this.partitionateListData(sortedListData);
    } else {
      partitionedListData = this.partitionateListData(listData);
    }


    if (partitionedListData.length !== 0) {



      // Fetching distinct users, so we are not fetching duplicates. Reducing number of requests.
      for (let y = 0; y < partitionedListData[this.state.currentPage].length; y++) {

        let userLoginName = partitionedListData[this.state.currentPage][y][strings.SchemaName].ValueString03;

        //if (!this.propertyExist(distinctValues, userLoginName)){
        distinctValues.push(userLoginName);

        userProfileService = new UserProfileService(this.props.context, userLoginName);
        promises.push(userProfileService.getUserProfileProperties());
        //}
      }


      Promise.all(promises).then((values) => {

        for (let x = 0; x < partitionedListData[this.state.currentPage].length; x++) {
          let userLoginName = partitionedListData[this.state.currentPage][x][strings.SchemaName].ValueString03;
          let userData;

          for (let z = 0; z < values.length; z++) {

            // check if no userprofile
            if (!values[z].hasOwnProperty('UserProfileProperties') || !values[z].UserProfileProperties) {
              continue;
            }

            if (getUserProfileProperty(values[z].UserProfileProperties, 'UserName') === userLoginName) {
              for (let y = 0; y < values[z].UserProfileProperties.length; y++) {
                if (values[z].UserProfileProperties[y].Key === 'PictureURL') {
                  let imageUrl = this.props.context.pageContext.site.absoluteUrl + '/_layouts/15/userphoto.aspx?size=M&username=' + userLoginName;
                  values[z].UserProfileProperties[y].Value = imageUrl;
                  break;
                }
              }
              userData = values[z];
              break;
            }
          }

          partitionedListData[this.state.currentPage][x].userProfile = userData;
        }

        this.setState({
          listData: partitionedListData,
          unPartitionedListData: listData,
          filter: groupedFilters,
          currentPage: 0,
          isLoading: false,
          navPages: navPages,
          emptyResult: false
        }, () => this.getSiteUrlsAsync(partitionedListData, listData, 0));
      });

    } else {
      this.setState({
        listData: partitionedListData,
        unPartitionedListData: listData,
        filter: groupedFilters,
        currentPage: 0,
        isLoading: false,
        navPages: navPages,
        emptyResult: true
      });
    }

    


  }

  public getSiteUrlsAsync = (partitionedListData, unPartitionedListData, currentPage) => {
    let promises: Promise<IGroupSiteUrl>[] = [];
    let groupSiteUrlService: GroupSiteUrlService;

    // Get site URL for all sites on current page.
    for (let x = 0; x < partitionedListData[currentPage].length; x++) {
      let id = partitionedListData[currentPage][x].id;
      groupSiteUrlService = new GroupSiteUrlService(this.props.context, id);
      promises.push(groupSiteUrlService.getGroupSiteUrl());
    }

    Promise.all(promises).then((values) => {

      for (let y = 0; y < values.length; y++) {
        if (values[y].displayName === partitionedListData[currentPage][y].displayName) {
          partitionedListData[currentPage][y]['siteUrl'] = values[y].webUrl;
          unPartitionedListData[y]['siteUrl'] = values[y].webUrl;
        } else {
          // match is not synced look for correct match.
          for (let z = 0; z < partitionedListData[currentPage].length; z++) {
            if (partitionedListData[currentPage][z].displayName === values[y].displayName) {
              partitionedListData[currentPage][z]['siteUrl'] = values[y].webUrl;
              unPartitionedListData[z]['siteUrl'] = values[y].webUrl;
            }
          }
        }
      }

      this.setState({
        listData: partitionedListData,
        unPartitionedListData: unPartitionedListData,
      });

    });


  }

  public partitionateListData(listData: IOffice365Group[]): Array<IOffice365Group[]> {
    let partitionatedData = [];
    const rowLimit = parseInt(this.props.listRows);
    let from = 0;
    let to = rowLimit;
    let numberOfIterationsNeeded = Math.ceil(listData.length / rowLimit);

    for (let partitionNumber = 0; partitionNumber < numberOfIterationsNeeded; partitionNumber++) {

      partitionatedData[partitionNumber] = listData.slice(from, to);
      from = to;
      to += rowLimit;

    }

    return partitionatedData;
  }

  public refreshListOnFilter = (filterData: Array<any>) => {
    let filter = filterData.map(f => { return { fieldName: f.fieldName, filterValue: f.filterValue }; });

    this.setState({
      filter: filter,
      navPages: [],
      listData: [],
      unPartitionedListData: [],
      currentPage: 0
    }, this.loadSiteList);
  }

  public buildFilter = (): string => {
    let filter = '';

    return filter;
  }

  public handleNewClick = (e) => {
    e.target.dispatchEvent(new CustomEvent('createSiteButtonClicked', { bubbles: true, detail: {} }));
  }

}
