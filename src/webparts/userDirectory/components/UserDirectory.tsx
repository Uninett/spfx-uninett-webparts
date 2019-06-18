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
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn, ConstrainMode, buildColumns, IDetailsRowProps, IDetailsRowStyles, DetailsRow } from 'office-ui-fabric-react/lib/DetailsList';
import { ShimmeredDetailsList } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';
import { mergeStyleSets, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Image, ImageFit, IImageStyles } from 'office-ui-fabric-react/lib/Image';
import { autobind, IconBase } from 'office-ui-fabric-react';

import { getTheme } from 'office-ui-fabric-react/lib/Styling';

const theme = getTheme();

const classNames = mergeStyleSets({
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap'
  },
  row: {
    
  }
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '300px'
  }
};


export default class UserDirectory extends React.Component<IUserDirectoryProps, IUserDirectoryState> {
 
  private _allUsers: IUserItem[];
  private _visibleColumns: IColumn[];
  private _showSorted: boolean;

  private _showSearchBox: boolean;
  private _searchBox: any;
  
  constructor(props: IUserDirectoryProps, state: IUserDirectoryState) {
    super(props);
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

    if (this.state.users === []) {
      // Render loading state ...

    } else {
      // Render real UI ...

    }

    return (
      <Fabric>
        <div className={classNames.controlWrapper}>
          <TextField label={strings.SearchBoxLabel} onChange={this._onChangeText} styles={controlStyles} />
        </div>

        <ShimmeredDetailsList
          items={users}
          compact={this.props.compactMode}
          columns={this._visibleColumns}
          selectionMode={SelectionMode.none}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          constrainMode={ConstrainMode.horizontalConstrained}
          onRenderItemColumn={_renderItemColumn}
          onRenderRow={this._onRenderRow}          
          enableShimmer={users.length == 0}
        />
      </Fabric>
    );
  }
  
  private _onRenderRow = (props: IDetailsRowProps): JSX.Element => {
    const customStyles: Partial<IDetailsRowStyles> = {};

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

    cols.push({
      key: 'colName',
      name: this.props.colNameTitle,
      fieldName: 'displayName',
      minWidth: 180,
      maxWidth: 200,
      //isRowHeader: true,
      //isSorted: true,
      //isSortedDescending: false,
      sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A',
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showJobTitle)
    cols.push({
      key: 'colTitle',
      name: this.props.colJobTitleTitle,
      fieldName: 'jobTitle',
      minWidth: 80,
      maxWidth: 90,
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showDepartment)
    cols.push({
      key: 'colDepartment',
      name: this.props.colDepartmentTitle,
      fieldName: 'department',
      minWidth: 100,
      maxWidth: 110,
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showOfficeLocation)
    cols.push({
      key: 'colOfficeLocation',
      name: this.props.colOfficeLocationTitle,
      fieldName: 'officeLocation',
      minWidth: 80,
      maxWidth: 90,
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showCity)
    cols.push({
      key: 'colCity',
      name: this.props.colCityTitle,
      fieldName: 'city',
      minWidth: 80,
      maxWidth: 90,
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showPhone)
    cols.push({
      key: 'colPhone',
      name: this.props.colPhoneTitle,
      fieldName: 'mobilePhone',
      minWidth: 90,
      maxWidth: 100,        
      data: 'string',      
      isPadded: true
    });

    if (this.props.showMail)
    cols.push({
      key: 'colMail',
      name: this.props.colMailTitle,
      fieldName: 'mail',
      minWidth: 240,
      maxWidth: 260,
      isCollapsible: false,
      data: 'string',
      onColumnClick: this._onColumnClick,
      isPadded: true
    });

    return cols;
  }

  private _onChangeText = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ): void => {
    // Filter by name or department (if not null)
    this.setState({
      users: text
        ? this._allUsers.filter(i => (i.displayName.toLowerCase().indexOf(text) > -1) || (i.department != null && i.department.toLowerCase().indexOf(text) > -1))
        : this._allUsers
    });
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

  @autobind
  private _search(): void {

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
      const customStyles: Partial<IImageStyles> = {};
      customStyles.image = {      
        borderRadius: 30/2
      };
      return <Image src={"/_layouts/15/userphoto.aspx?size=L&username=" + fieldContent} width={30} height={30} imageFit={ImageFit.cover} shouldFadeIn={true} styles={customStyles}/>;

    case 'colMail':
      return <Link href={"mailto:" + fieldContent}>{fieldContent}</Link>;

    case 'colPhone':
      return <Link href={"tel:" + fieldContent}>{fieldContent}</Link>;

    default:
      return <span>{fieldContent}</span>;
  }
}