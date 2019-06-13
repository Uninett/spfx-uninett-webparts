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
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn, ConstrainMode, buildColumns } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { mergeStyleSets, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { autobind } from 'office-ui-fabric-react';

const classNames = mergeStyleSets({
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap'
  },
  exampleToggle: {
    display: 'inline-block',
    marginBottom: '10px',
    marginRight: '30px'
  },
  selectionDetails: {
    marginBottom: '20px'
  }
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '300px'
  }
};


export default class UserDirectory extends React.Component<IUserDirectoryProps, IUserDirectoryState> {

  private _selection: Selection;  
  private _allUsers: IUserItem[];
  
  constructor(props: IUserDirectoryProps, state: IUserDirectoryState) {
    super(props);
    this._search();

    //console.log("Number of users (start of constructor): " + this.state.users);

    const columns: IColumn[] = [
      {
        key: 'colPhoto',
        name: '',
        fieldName: 'photo',
        minWidth: 30,
        maxWidth: 30,
        //isPadded: true
      },
      {
        key: 'colName',
        name: 'Name',
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
      },
      {
        key: 'colTitle',
        name: 'Job Title',
        fieldName: 'jobTitle',
        minWidth: 70,
        maxWidth: 90,
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      },
      {
        key: 'colMail',
        name: 'Mail',
        fieldName: 'mail',
        minWidth: 240,
        maxWidth: 260,
        isCollapsible: false,
        data: 'string',
        onColumnClick: this._onColumnClick,
        isPadded: true
      },
      {
        key: 'colPhone',
        name: 'Phone',
        fieldName: 'mobilePhone',
        minWidth: 70,
        maxWidth: 90,        
        data: 'string',
        //onColumnClick: this._onColumnClick
      }
    ];

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails()
        });
      }
    });
    
    this.state = {
      users: [],
      columns: columns,
      selectionDetails: this._getSelectionDetails(),
      isModalSelection: false,
      isCompactMode: this.props.compactMode
    };

  }

  public render(): React.ReactElement<IUserDirectoryProps> {
    
    let cols: IColumn[] = [];
    
    if (this.props.showPhoto)
    cols.push({
      key: 'colPhoto',
      name: '',
      fieldName: 'photo',
      minWidth: 30,
      maxWidth: 30,
    });

    cols.push({
      key: 'colName',
      name: 'Name',
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
      name: 'Job Title',
      fieldName: 'jobTitle',
      minWidth: 70,
      maxWidth: 90,
      onColumnClick: this._onColumnClick,
      data: 'string',
      isPadded: true
    });

    if (this.props.showMail)
    cols.push({
      key: 'colMail',
      name: 'Mail',
      fieldName: 'mail',
      minWidth: 240,
      maxWidth: 260,
      isCollapsible: false,
      data: 'string',
      onColumnClick: this._onColumnClick,
      isPadded: true
    });

    if (this.props.showPhone)
    cols.push({
      key: 'colPhone',
      name: 'Phone',
      fieldName: 'mobilePhone',
      minWidth: 70,
      maxWidth: 90,        
      data: 'string'
    });
    

    const { columns, isCompactMode, users, selectionDetails, isModalSelection } = this.state;
    return (
      <Fabric>
        {/*
        <div className={classNames.controlWrapper}>
          <Toggle
            label="Enable compact mode"
            checked={isCompactMode}
            onChange={this._onChangeCompactMode}
            onText="Compact"
            offText="Normal"
            styles={controlStyles}
          />
          <Toggle
            label="Enable modal selection"
            checked={isModalSelection}
            onChange={this._onChangeModalSelection}
            onText="Modal"
            offText="Normal"
            styles={controlStyles}
          />
          <TextField label="Filter by name:" onChange={this._onChangeText} styles={controlStyles} />
        </div>
        <div className={classNames.selectionDetails}>{selectionDetails}</div>
        */}

        <div className={classNames.controlWrapper}>
          <TextField label="Filter by name:" onChange={this._onChangeText} styles={controlStyles} />
        </div>

        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={users}
            compact={this.props.compactMode}
            columns={cols}
            selectionMode={isModalSelection ? SelectionMode.multiple : SelectionMode.none}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selection={this._selection}
            selectionPreservedOnEmptyClick={false}
            //onItemInvoked={this._onItemInvoked}
            enterModalSelectionOnTouch={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            constrainMode={ConstrainMode.unconstrained}
            onRenderItemColumn={_renderItemColumn}
          />
        </MarqueeSelection>
      </Fabric>
    );
  }
  public componentDidUpdate(previousProps: any, previousState: IUserDirectoryState) {
    if (
      previousState.isModalSelection !== this.state.isModalSelection &&
      !this.state.isModalSelection
    ) {
      this._selection.setAllSelected(false);
    }
  }

  private _onChangeCompactMode = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
    this.setState({ isCompactMode: checked });
  }

  private _onChangeModalSelection = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
    this.setState({ isModalSelection: checked });
  }

  private _onChangeText = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ): void => {
    this.setState({
      users: text
        ? this._allUsers.filter(i => i.displayName.toLowerCase().indexOf(text) > -1)
        : this._allUsers
    });
  }

  private _onItemInvoked(item: any): void {
    alert(`Item invoked: ${item.name}`);
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IUserItem).displayName;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, users } = this.state;
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
    this.setState({
      columns: newColumns,
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
          .select("userPrincipalName,displayName,jobTitle,mail,mobilePhone")
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
                mail: item.mail,
                mobilePhone: item.mobilePhone
              });
            });
            
            // Update the component state accordingly to the result
            this.setState(
              {
                users: users,
              }
            );
            console.log("Number of users (end of client): " + users.length);
            
            // Update all-users array and re-render component
            this._allUsers = users;
            this.render();
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
      return <Image src={"/_layouts/15/userphoto.aspx?size=L&username=" + fieldContent} width={30} height={30} imageFit={ImageFit.cover} />;

    case 'colMail':
      return <Link href={"mailto:" + fieldContent}>{fieldContent}</Link>;

    case 'colPhone':
      return <Link href={"tel:" + fieldContent}>{fieldContent}</Link>;

    default:
      return <span>{fieldContent}</span>;
  }
}