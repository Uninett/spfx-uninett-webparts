import * as React from "react";
import { DefaultButton, CompoundButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { PeoplePicker } from './PeoplePicker';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ParentDepartment } from './ParentDepartment';
import { Checkbox, ICheckboxStyles, ICheckboxProps } from 'office-ui-fabric-react/lib/Checkbox';
import { IDepartmentState } from './IDepartmentState';
import { SharePointUserPersona } from '../models/PeoplePicker';
import * as strings from 'OrderGroupWebPartStrings';
import styles from './OrderGroup.module.scss';
import { PrivacySetting } from './PrivacySetting';
import { SiteType } from '../interfaces/SiteType';

export interface IDepartmentProps {
    cancel: () => void;
    context: WebPartContext;
    updateList: (data: any) => void;
    departmentSiteTypeName: SiteType;
    hideParentDepartment: boolean;
}

class Department extends React.Component<IDepartmentProps, any> {

    constructor(props: any) {
        super(props);
        this.state = {
            departmentName: "",
            departmentDescription: "",
            // shortName: "",
            externalShare: false,
            createTeam: true,
            privacySetting: "Closed",
            people: [],
            parentDepartment: ""
        };
    }

    public render() {
        let newDeparmentLabel = this.props.departmentSiteTypeName === 0 ? strings.NewDepartment : strings.NewSection;

        return (<div>
            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <label>{newDeparmentLabel}</label>
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <Label required={true}>{strings.DepartmentLeader}</Label>
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
                            this.setState({ people });
                        }}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <TextField
                        label={strings.DepartmentName}
                        maxLength={255}
                        onChange={(_, text: string) => this.setState({ departmentName: text })}
                        required={true}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <TextField
                        label={strings.Description}
                        maxLength={500}
                        onChange={(_, text: string) => this.setState({ departmentDescription: text })}
                    />
                </div>
            </div>

            {/* <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <TextField
                        label={strings.ShortName}
                        maxLength={20}
                        onChange={(text: string) => this.setState({ shortName: text })}
                    />
                </div>
            </div> */}

            {!this.props.hideParentDepartment ? (
                <div className="ms-Grid-row">
                    <div className={styles.positioning}>
                        <ParentDepartment
                            context={this.props.context}
                            onChange={(option) => {
                                this.setState({ parentDepartment: option.text });
                            }}
                        />
                    </div>
                </div>
            ) : ''}

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <PrivacySetting
                    onChange={(option) => {
                        this.setState({ privacySetting: option.key });
                    }}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <Checkbox
                        label={strings.CreateMicrosoftTeam}
                        id='checkbox2'
                        defaultChecked={true}
                        onChange={this._onCreateTeamChange}
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

        </div>);
    }

    private _onCreateTeamChange = (ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {
      if (isChecked == false) {
          this.setState({ createTeam: false });
      }
      else {
          this.setState({ createTeam: true });
      }
    }

    private _onFinishClick = () => {

        if (this.state.people.length > 0 && this.state.departmentName != "") {

            if (!this.props.hideParentDepartment && this.state.parentDepartment == "") {
                alert(strings.IncompleteAlert);
            } else {
                let data = {
                    'ContentTypeId': '0x01006723DC5874B448DBA9EB9966245656FA',
                    'KDTOOwnerId': this.state.people[0].User.Id,
                    'Title': this.state.departmentName,
                    'KDTOSiteDescription': this.state.departmentDescription,
                    // 'KDTOShortName': this.state.shortName,
                    'KDTOParentDepartment': this.state.parentDepartment,
                    'KDTOSitePrivacy': this.state.privacySetting,
                    'KDTOExternalSharing': this.state.externalShare,
                    'KDTOCreateTeam': this.state.createTeam
                };
                this.props.updateList(data);
            }

        }
        else {
            alert(strings.IncompleteAlert);
        }
    }
}

export { Department };
