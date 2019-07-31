import * as React from 'react';
import { DefaultButton, CompoundButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { PeoplePicker } from './PeoplePicker';
import { IActivityState } from './IActivityState';
import { IWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';
import { ParentDepartment } from './ParentDepartment';
import { Checkbox, ICheckboxStyles, ICheckboxProps } from 'office-ui-fabric-react/lib/Checkbox';
import { SharePointUserPersona } from '../models/PeoplePicker';
import * as strings from 'OrderGroupWebPartStrings';
import styles from './OrderGroup.module.scss';
import { PrivacySetting } from './PrivacySetting';

export interface IActivityProps {
    cancel: () => void;
    context: WebPartContext;
    updateList: (data: any) => void;
    hideParentDepartment: boolean;
}

class Activity extends React.Component<IActivityProps, IActivityState> {

    constructor(props: any) {
        super(props);
        this.state = {
            activityName: "",
            activityDescription: "",
            parentDepartment: "",
            privacySetting: "Closed",
            externalShare: true,
            people: []
        };
    }

    render() {
        return (<div>
            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <Label>{strings.NewActivity}</Label>
                </div>
            </div>


            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <Label required={true}>{strings.ActivityResponsible}</Label>
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
                        onChange={(people: SharePointUserPersona[]) => {
                            this.setState({ people })
                            var emails = people.map(spPersona => {
                                return spPersona.User.Email
                            })
                        }}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <TextField
                        label={strings.ActivityName}
                        maxLength={255}
                        onChange={(_, text: string) => this.setState({ activityName: text })}
                        required={true}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <TextField
                        label={strings.ActivityDescription}
                        maxLength={500}
                        onChange={(_, text: string) => this.setState({ activityDescription: text })}
                    />
                </div>
            </div>
            
            {!this.props.hideParentDepartment ? (
                <div className="ms-Grid-row">
                    <div className={styles.positioning}>
                        <ParentDepartment
                            context={this.props.context}
                            onChange={(option) => {
                                this.setState({ parentDepartment: option.text })
                            }}
                        />
                    </div>
                </div>
            ) : ''}
            

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <PrivacySetting
                    onChange={(option) => {
                        this.setState({ privacySetting: option.key })
                    }}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <Checkbox
                        label={strings.OpenForExternalSharing}
                        id='checkbox1'
                        defaultChecked={true}
                        onChange={this._onCheckboxChange}
                    />
                </div>
            </div>


            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <DefaultButton
                        text={strings.Cancel}
                        onClick={this.props.cancel}
                    />

                    <DefaultButton
                        primary={true}
                        text={strings.Finish}
                        onClick={this._onFinishClick}
                    />
                </div>
            </div>

        </div>)
    }

    private _onCheckboxChange = (ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {
        if (isChecked == false) {
            this.setState({ externalShare: false });
        }
        else {
            this.setState({ externalShare: true });
        }
    }

    private _onFinishClick = () => {
        if (this.state.people.length > 0 && this.state.activityName !== "") {
            
            if (!this.props.hideParentDepartment && this.state.parentDepartment == "") {
                alert(strings.IncompleteAlert);
            } else {
                let data = {
                    'ContentTypeId': '0x0100DB027F48E8544E67AE646920383C7DC6',
                    'KDTOOwnerId': this.state.people[0].User.Id,
                    'Title': this.state.activityName,
                    'KDTOSiteDescription': this.state.activityDescription,
                    'KDTOParentDepartment': this.state.parentDepartment,
                    'KDTOSitePrivacy': this.state.privacySetting,
                    'KDTOExternalSharing': this.state.externalShare
                };
                this.props.updateList(data);
            }
        }
        else {
            alert(strings.IncompleteAlert);
        }
    }

}

export { Activity }