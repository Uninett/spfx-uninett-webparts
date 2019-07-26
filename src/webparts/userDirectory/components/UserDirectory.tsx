import * as React from 'react';
import styles from './UserDirectory.module.scss';
import * as strings from 'UserDirectoryWebPartStrings';
import { IUserDirectoryProps } from './IUserDirectoryProps';
import { IUserDirectoryState } from './IUserDirectoryState';
import { IUserItem } from './IUserItem';

import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from "@microsoft/sp-http";

import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn, ConstrainMode, buildColumns, IDetailsRowProps, IDetailsRowStyles, DetailsRow} from 'office-ui-fabric-react/lib/DetailsList';
import { ShimmeredDetailsList } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Image, ImageFit, IImageStyles } from 'office-ui-fabric-react/lib/Image';
import { IconBase, Stack, SearchBox, ITooltipHostProps } from 'office-ui-fabric-react';
import { getTheme } from 'office-ui-fabric-react/lib/Styling';

import { RxJsEventEmitter } from '../../../RxJsEventEmitter/RxJsEventEmitter';
import IEventData from '../../../RxJsEventEmitter/IEventData';


const theme = getTheme();

const searchBoxStyle = {
  root: {
    margin: '0 30px 20px 0',
    width: 300 
  }
};

const photoSize = 30;

const imageStyle = {
  image: {
    borderRadius: photoSize/2
  }
};


export default class UserDirectory extends React.Component<IUserDirectoryProps, IUserDirectoryState> {
  private readonly eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  
  private _allUsers: IUserItem[];
  private _visibleColumns: IColumn[];
  private _showSorted: boolean;
  private _isApiCorrect: boolean;
  
  constructor(props: IUserDirectoryProps, state: IUserDirectoryState) {
    super(props);
    this.eventEmitter.on("filterTerms", this._onReceiveData.bind(this));    
    this._isApiCorrect = true;
    this._search();
    this._showSorted = false;

    this.state = {
      users: []
    };
  } 

  public render(): React.ReactElement<IUserDirectoryProps> {

    const { users } = this.state;

    // Check flag to avoid rendering un-sorted columns
    if (!this._showSorted) {
      this._visibleColumns = this._getCheckedColumns();
    }
    this._showSorted = false;
    
    // Display error message if API not valid
    if (!this._isApiCorrect) {
      return (        
        <Fabric>
          { this.props.useBuiltInSearch && (
            <SearchBox
              styles={searchBoxStyle}
              placeholder={this.props.searchBoxPlaceholder}
              onChange={(_, newValue) => this._onChangeText(newValue)}       
            />
          )}          
          <ShimmeredDetailsList
            items={users}
            columns={this._visibleColumns}
            selectionMode={SelectionMode.none}
          />
          <Stack horizontalAlign='center'>
            <Text block>{strings.BadApi1}</Text>
            <Text block>{strings.BadApi2}</Text>
          </Stack>
      </Fabric>
      );
    }

    return (
      <Fabric>
        { this.props.useBuiltInSearch && (
            <SearchBox
              styles={searchBoxStyle}
              placeholder={this.props.searchBoxPlaceholder}
              onChange={(_, newValue) => this._onChangeText(newValue)}
            />
          )}
        <ShimmeredDetailsList
          items={users}
          compact={this.props.compactMode}
          columns={this._visibleColumns}
          selectionMode={SelectionMode.none}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          constrainMode={ConstrainMode.horizontalConstrained}
          //onRenderDetailsHeader={this._onRenderDetailsHeader}
          onRenderItemColumn={_renderItemColumn}
          onRenderRow={this._onRenderRow}          
          enableShimmer={this._allUsers == undefined}
          shimmerLines={30}
        />
        { users.length == 0 && (
         <Stack horizontalAlign='center'>
          <Text>{strings.NoUsers}</Text>
         </Stack>
        )}
      </Fabric>
    );
  }

  // Attempt to make column headers bold
  /*
  private _onRenderDetailsHeader(detailsHeaderProps: IDetailsHeaderProps) {
    const customStyles: Partial<IDetailsHeaderStyles> = {};

    customStyles.root = {
      //fontSize: "14px",
      fontWeight: 600
    };
    console.log("style set");
    return (
      <DetailsHeader
        {...detailsHeaderProps}
        styles={customStyles}
      />
    );
  }
  */

  private _onRenderRow = (props: IDetailsRowProps): JSX.Element => {
    const customStyles: Partial<IDetailsRowStyles> = {};

    // Vertically center content of each row
    customStyles.fields = {      
      alignItems: "center"
    };

    if (this.props.alternatingColours) {
      if (props.itemIndex % 2 === 0) {
        // Every other row renders with a different background color
        customStyles.fields = { 
          backgroundColor: theme.palette.themeLighterAlt,
          alignItems: "center"
          };
      }
    }
    return <DetailsRow {...props} styles={customStyles} />;
  }

  private _onReceiveData = (
    data: IEventData
  ): void => {
    if (this._allUsers != undefined) {
      // Filter by name or department (if not null)
      let text: string = data.sharedData.toLowerCase();
      this.setState({
        users: text
          ? this._allUsers.filter(i => (i.displayName.toLowerCase().indexOf(text) > -1) || (i.department != null && i.department.toLowerCase().indexOf(text) > -1))
          : this._allUsers
      });
    }
  }

  private _onChangeText = (
    text: string
  ): void => {
    if (this._allUsers != undefined) {
      // Filter by name or department (if not null)
      text = text.toLowerCase();
      this.setState({
        users: text
          ? this._allUsers.filter(i => (i.displayName.toLowerCase().indexOf(text) > -1) || (i.department != null && i.department.toLowerCase().indexOf(text) > -1))
          : this._allUsers
      });
    }
  }

  private _getCheckedColumns(): IColumn[] {
    // Return an array of columns that the user has selected in the web part's property pane
    let cols: IColumn[] = [];
    
    if (this.props.showPhoto)
    cols.push({
      key: 'colPhoto',
      name: '',
      fieldName: 'userPrincipalName',
      minWidth: 30,
      maxWidth: 30,
    });

    if (this.props.showName)
    cols.push({
      key: 'colName',
      name: this.props.colName,
      fieldName: 'displayName',
      minWidth: 180,
      maxWidth: 200,
      isResizable: true,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showJobTitle)
    cols.push({
      key: 'colJobTitle',
      name: this.props.colJobTitle,
      fieldName: 'jobTitle',
      minWidth: 120,
      maxWidth: 140,
      isResizable: true,
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showDepartment)
    cols.push({
      key: 'colDepartment',
      name: this.props.colDepartment,
      fieldName: 'department',
      minWidth: 120,
      maxWidth: 140,
      isResizable: true,
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showOfficeLocation)
    cols.push({
      key: 'colOfficeLocation',
      name: this.props.colOfficeLocation,
      fieldName: 'officeLocation',
      minWidth: 80,
      maxWidth: 120,
      isResizable: true,
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showCity)
    cols.push({
      key: 'colCity',
      name: this.props.colCity,
      fieldName: 'city',
      minWidth: 80,
      maxWidth: 120,
      isResizable: true,
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showPhone)
    cols.push({
      key: 'colPhone',
      name: this.props.colPhone,
      fieldName: 'mobilePhone',
      minWidth: 90,
      maxWidth: 100,       
      data: 'string',      
      isPadded: true
    });

    if (this.props.showMail)
    cols.push({
      key: 'colMail',
      name: this.props.colMail,
      fieldName: 'mail',
      minWidth: 170,
      maxWidth: 250,
      isResizable: true,
      isCollapsible: false,
      data: 'string',
      onColumnClick: this._onColumnClick,
      isPadded: true
    });

    return cols;
  }

  

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {

    const columns = this._visibleColumns;
    const users = this.state.users;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newUsers = _copyAndSort(users, currColumn.fieldName!, currColumn.isSortedDescending);
    // Set flag to avoid rendering un-sorted columns
    this._showSorted = true;
    this._visibleColumns = newColumns;
    this.setState({
      users: newUsers
    });
  }

  private _search = (): void => {
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        // From https://github.com/microsoftgraph/msgraph-sdk-javascript sample
        client
          .api(this.props.api)
          .version("v1.0")
          .select("userPrincipalName,displayName,jobTitle,department,officeLocation,city,mobilePhone,mail")
          .get((err, res) => {
  
            if (err) {
              console.error(err);
              // Re-renders component with error message
              this._isApiCorrect = false;
              this.forceUpdate();
              return;
            }            
            
            // Prepare the output array
            var users: Array<IUserItem> = new Array<IUserItem>();

            // Map the JSON response to the output array
            res.value.map((item: any) => {
              users.push( { 
                userPrincipalName: item.userPrincipalName,
                displayName: item.displayName,
                jobTitle: item.jobTitle,
                department: item.department,
                officeLocation: item.officeLocation,
                city: item.city,            
                mobilePhone: item.mobilePhone,
                mail: item.mail
              });
            });
                        
            // Update user array and component state
            this._allUsers = users;
            this.setState(
              {
                users: users,
              }
            );
            console.log("Search completed");
          });
      });

      
  }  
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items
    .slice(0)
    .sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}

function _renderItemColumn(item: IUserItem, index: number, column: IColumn) {
  const fieldContent = item[column.fieldName as keyof IUserItem] as string;

  switch (column.key) {
    case 'colPhoto':      
      return <Image src={"/_layouts/15/userphoto.aspx?size=L&username=" + fieldContent} width={photoSize} height={photoSize} imageFit={ImageFit.cover} shouldFadeIn={false} styles={imageStyle}/>;

    case 'colMail':
      return <Link href={"mailto:" + fieldContent}>{fieldContent}</Link>;

    case 'colPhone':
      return <Link href={"tel:" + fieldContent}>{fieldContent}</Link>;

    default:
      return <span>{fieldContent}</span>;
  }
}