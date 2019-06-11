import * as React from 'react';
import styles from './UserDirectory.module.scss';
import * as strings from 'UserDirectoryWebPartStrings';
import { IUserDirectoryProps } from './IUserDirectoryProps';
import { IUserDirectoryState } from './IUserDirectoryState';
import { IUserItem } from './IUserItem';

import { escape } from '@microsoft/sp-lodash-subset';
import { 
  ListView, 
  IViewField, 
  SelectionMode, 
  GroupOrder, 
  IGrouping 
} from "@pnp/spfx-controls-react/lib/ListView";

import {
  autobind,
  PrimaryButton,
  TextField,
  Label,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility
} from 'office-ui-fabric-react';

import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";

// Configure the columns for the listView component
let _listViewColumns = [
  {
    name: 'displayName',
    displayName: 'Display name',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    name: 'mail',
    displayName: 'Mail',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    name: 'userPrincipalName',
    displayName: 'User Principal Name',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  }
];

export default class UserDirectory extends React.Component<IUserDirectoryProps, IUserDirectoryState> {

  constructor(props: IUserDirectoryProps, state: IUserDirectoryState) {
    super(props);

    // Initialize the state of the component
    this.state = {
      users: [],
      searchFor: ""
    };
  }

  public render(): React.ReactElement<IUserDirectoryProps> {

    return (
      <div>
        <span className={ styles.title }>Search for a user!</span>
        <p className={ styles.form }>
          <TextField 
              label={ strings.SearchFor } 
              required={ true } 
              value={ this.state.searchFor }
              onChanged={ this._onSearchForChanged }
              onGetErrorMessage={ this._getSearchForErrorMessage }
            />
        </p>
        <p className={styles.form}>
              <PrimaryButton
                text='Search'
                title='Search'
                onClick={this._search}
              />
        </p>
        
        {
          (this.state.users != null && this.state.users.length > 0) ?
            <p className={ styles.form }>
            {
              <ListView
              items={ this.state.users }
              viewFields={_listViewColumns}
              compact={true}
              selectionMode={SelectionMode.single}/>
            }
          </p>
          : null
        }
      </div>
    );
  }

  @autobind
  private _onSearchForChanged(newValue: string): void {

    // Update the component state accordingly to the current user's input
    this.setState({
      searchFor: newValue,
    });
  }

  private _getSearchForErrorMessage(value: string): string {
    // The search for text cannot contain spaces
    return (value == null || value.length == 0 || value.indexOf(" ") < 0)
      ? ''
      : `${strings.SearchForValidationErrorMessage}`;
  }

  @autobind
  private _search(): void {

    // Log the current operation
    console.log("Using _search() method");
    
    var filter = "";
    if (this.state.searchFor != "") {
      filter = `startswith(displayName, '${escape(this.state.searchFor)}') or startswith(givenName, '${escape(this.state.searchFor)}') or startswith(surname, '${escape(this.state.searchFor)}')`;
    }

    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        // From https://github.com/microsoftgraph/msgraph-sdk-javascript sample
        client
          .api(this.props.api)
          .version("v1.0")
          .select("displayName,mail,userPrincipalName")
          .filter(filter)
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
                displayName: item.displayName,
                mail: item.mail,
                userPrincipalName: item.userPrincipalName,
              });
            });
  
            // Update the component state accordingly to the result
            this.setState(
              {
                users: users,
              }
            );
          });
      });
  }

}
