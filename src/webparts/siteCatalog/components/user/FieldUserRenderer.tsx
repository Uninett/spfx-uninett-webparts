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

export interface IFieldUserRendererProps extends IFieldRendererProps {
    /**
     * users to be displayed
     */
    //users?: IPrincipal[];
    users?: IFieldUser[];
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

export interface IFieldUserRendererState {
    users?: IFieldUser[];
}

/**
 * Field User Renderer.
 * Used for:
 *   - People and Groups
 */
export class FieldUserRenderer extends React.Component<IFieldUserRendererProps, IFieldUserRendererState> {

    // cached user profiles
    //private _loadedUserProfiles: { [id: string]: IUserProfileProperties } = {};
    private _userUrlTemplate: string;
    private _userImageUrl: string;


    public constructor(props: IFieldUserRendererProps, state: IFieldUserRendererState) {
        super(props, state);

        this.state = {
            users: []
        };
        
    }

    /*componentWillMount() {
        console.log('inside componentWillMount');
        if (this.props.group.hasOwnProperty('userProfile')) {
            let user:IFieldUser = {
                id: this.props.group.userProfile.UserProfileProperties[0].Value,
                email: this.props.group.userProfile.Email,
                currentUser: false,
                displayName: this.props.group.userProfile.DisplayName,
                imageUrl: this.props.group.userProfile.UserProfileProperties[18].Value,
                userUrl: this.props.group.userProfile.UserUrl
            }

            let col = [user];

            this.setState({
                users: col
            });
        } else {
            console.log('group has no user profile: ');
            console.log(this.props.group);
            console.log('');
        }
    }*/

    //componentWillReceiveProps(nextProps: IFieldUserRendererProps){
        /*console.log('inside componentwillreceiceprops');
        let user:IFieldUser;
        let col;
        if (this.props.group.hasOwnProperty('userProfile')) {
            user = {
                id: this.props.group.userProfile.UserProfileProperties[0].Value,
                email: this.props.group.userProfile.Email,
                currentUser: false,
                displayName: this.props.group.userProfile.DisplayName,
                imageUrl: this.props.group.userProfile.UserProfileProperties[18].Value,
                userUrl: this.props.group.userProfile.UserUrl
            }

            col = [user];
        }

        if (
          col &&
          !deepEqual(col, this.state.users)
        ){
          this.setState({
              users: col
          });
        }*/
      //}

    @override
    public render(): JSX.Element {
        const userEls: JSX.Element[] = this.props.users.map((user, index) => {
            const expandingCardProps: IExpandingCardProps = {
                onRenderCompactCard: (user.email ? this._onRenderCompactCard.bind(this, index) : null),
                onRenderExpandedCard: (user.email ? this._onRenderExpandedCard.bind(this) : null),
                renderData: user,
                directionalHint: DirectionalHint.bottomLeftEdge,
                gapSpace: 0,
                expandedCardHeight: 150,
                trapFocus: true
            };
            const hoverCardProps: IFieldUserHoverCardProps = {
                expandingCardProps: expandingCardProps,
                displayName: user.displayName,
                cssProps: this.props.cssProps
            };
            return <FieldUserHoverCard {...hoverCardProps} />;
        });
        return <div style={this.props.cssProps} className={css(this.props.className)}>{userEls}</div>;
    }

    /**
     * Renders compact part of user Hover Card
     * @param index user index in the list of users/groups in the People and Group field value
     * @param user IUser
     */
    private _onRenderCompactCard(index: number, user: IFieldUser): JSX.Element {
        const sip: string = user.sip || user.email;
        let actionsEl: JSX.Element;
        if (user.currentUser) {
            actionsEl = <div className={styles.actions}>
                <Button buttonType={ButtonType.command} iconProps={{ iconName: 'Edit' }} href={user.userUrl} target={'_blank'}>{strings.UpdateProfile}</Button>
            </div>;
        }
        else {
            actionsEl = <div className={styles.actions}>
                <IconButton iconProps={{ iconName: 'Mail' }} title={strings.SendEmailTo.replace('{0}', user.email)} href={`mailto:${user.email}`} />
                <IconButton iconProps={{ iconName: 'Chat' }} title={strings.StartChatWith.replace('{0}', sip)} href={`sip:${sip}`} className={styles.chat} />
            </div>;
        }

        return <div className={styles.main}>
            <Persona
                imageUrl={user.imageUrl}
                primaryText={user.displayName}
                secondaryText={user.department}
                tertiaryText={user.jobTitle}
                size={PersonaSize.large} />
            {actionsEl}
        </div>;
    }

    /**
     * Renders expanded part of user Hover Card
     * @param user IUser
     */
    private _onRenderExpandedCard(user: IFieldUser): JSX.Element {
        //if (this._loadedUserProfiles[user.id]) {
            return <ul className={styles.sections}>
                <li className={styles.section}>
                    <div className={styles.header}>{strings.Contact} <i className={css('ms-Icon ms-Icon--ChevronRight', styles.chevron)} aria-hidden={'true'}></i></div>
                    <div className={styles.contactItem}>
                    <Icon iconName={'Mail'}/>
                        <Link className={styles.content} title={user.email} href={`mailto:${user.email}`} target={'_self'}>{user.email}</Link>
                    </div>
                    {user.workPhone &&
                        <div className={styles.contactItem}>
                            <Icon iconName={'Phone'}/>
                            <Link className={styles.content} title={user.workPhone} href={`tel:${user.workPhone}`} target={'_self'}>{user.workPhone}</Link>
                        </div>
                    }
                    {user.cellPhone &&
                        <div className={styles.contactItem}>
                            <Icon iconName={'Phone'}/>
                            <Link className={styles.content} title={user.cellPhone} href={`tel:${user.cellPhone}`} target={'_self'}>{user.cellPhone}</Link>
                        </div>
                    }
                </li>
            </ul>;
        /*}
        else {
            return <Spinner size={SpinnerSize.large} />;
        }*/
    }

}