import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';
import * as _ from '@microsoft/sp-lodash-subset';
import { HoverCard, IExpandingCardProps, DirectionalHint, Persona, IconButton, Button, ButtonType, PersonaSize, Spinner, SpinnerSize, Link, Icon } from 'office-ui-fabric-react';

import { IPrincipal } from '../interfaces/SPEntities';
import { IFieldRendererProps } from './IFieldRendererProps';

import styles from './FieldUserRenderer.module.scss';
//import { SPHttpClient } from '@microsoft/sp-http';
import FieldUserHoverCard, { IFieldUserHoverCardProps } from './FieldUserHoverCard';
//import * as appInsights from '../../../common/appInsights';

import * as strings from 'SiteCatalogWebPartStrings';

import { IOffice365Group } from '../interfaces/office365Group';
import { deepEqual } from '../helpers/deepEqual';
import { FieldUserRenderer } from './FieldUserRenderer';
import { getUserProfileProperty } from '../helpers/userProfile';

export interface IFieldUserContainerProps extends IFieldRendererProps {
    /**
     * users to be displayed
     */
    //users?: IPrincipal[];
    group: IOffice365Group;
}

/**
 * Internal interface to work with user profile
 */
export interface IFieldUser {
    /**
     * display  name
     */
    displayName?: string;
    /**
     * job title
     */
    jobTitle?: string;
    /**
     * department
     */
    department?: string;
    /**
     * user id
     */
    id?: string;
    /**
     * avatar url
     */
    imageUrl?: string;
    /**
     * email
     */
    email?: string;
    /**
     * skype for business username
     */
    sip?: string;
    /**
     * true if the user is current user
     */
    currentUser?: boolean;
    /**
     * work phone
     */
    workPhone?: string;
    /**
     * cell phone
     */
    cellPhone?: string;
    /**
     * url to edit user profile in Delve
     */
    userUrl?: string;
}

export interface IFieldUserContainerState {
    users?: IFieldUser[];
}

/**
 * Field User Renderer.
 * Used for:
 *   - People and Groups
 */
export class FieldUserContainer extends React.Component<IFieldUserContainerProps, IFieldUserContainerState> {

    // cached user profiles
    //private _loadedUserProfiles: { [id: string]: IUserProfileProperties } = {};
    private _userUrlTemplate: string;
    private _userImageUrl: string;


    public constructor(props: IFieldUserContainerProps, state: IFieldUserContainerState) {
        super(props, state);

        this.state = {
            users: []
        };
        
    }

    public componentWillMount() {
        
        if (this.props.group.hasOwnProperty('userProfile') && this.props.group['userProfile']) {
            let imageUrl = getUserProfileProperty(this.props.group.userProfile.UserProfileProperties, 'PictureURL');
            imageUrl = decodeURIComponent(imageUrl);

            let user:IFieldUser = {
                id: this.props.group.userProfile.UserProfileProperties[0].Value,
                email: this.props.group.userProfile.Email,
                currentUser: false,
                displayName: this.props.group.userProfile.DisplayName,
                imageUrl: imageUrl,
                userUrl: this.props.group.userProfile.UserUrl
            };

            let col = [user];

            this.setState({
                users: col
            });
        } else {
            /*console.log('group has no user profile: ');
            console.log(this.props.group);
            console.log('');*/
        }
    }

    public componentWillReceiveProps(nextProps: IFieldUserContainerProps){
        let user:IFieldUser;
        let col;

        if (this.props.group.hasOwnProperty('userProfile') && this.props.group['userProfile']) {
            let imageUrl = getUserProfileProperty(this.props.group.userProfile.UserProfileProperties, 'PictureURL');
            imageUrl = decodeURIComponent(imageUrl);
            user = {
                id: this.props.group.userProfile.UserProfileProperties[0].Value,
                email: this.props.group.userProfile.Email,
                currentUser: false,
                displayName: this.props.group.userProfile.DisplayName,
                imageUrl: imageUrl,
                userUrl: this.props.group.userProfile.UserUrl
            };

            col = [user];
        } else {
            //console.log('user profile not found');
        }

        if (
          col &&
          !deepEqual(col, this.state.users)
        ){
            console.log('updating user state from ', this.state.users);
            console.log('to: ', col);
          this.setState({
              users: col
          });
        } else {
            //console.log('skipping state update for group ', col);
        }
      }

    @override
    public render(): JSX.Element {
        return (
            <div>
                {this.state.users.length > 0 ? <FieldUserRenderer users={this.state.users} /> : this.props.group['extvcs569it_InmetaGenericSchema']['LabelString03']}
            </div>
        );
    }

}