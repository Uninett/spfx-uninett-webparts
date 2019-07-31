import * as React from "react";
import { DefaultButton, CompoundButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { PeoplePicker } from './PeoplePicker';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Checkbox, ICheckboxStyles, ICheckboxProps } from 'office-ui-fabric-react/lib/Checkbox';
import { ParentDepartment } from './ParentDepartment';
import { SharePointUserPersona } from '../models/PeoplePicker';
import TaxonomyPicker from "react-taxonomypicker";
import "react-taxonomypicker/dist/React.TaxonomyPicker.css";
import TaxonomyPickerLoader from '../components/TaxonomyPicker/TaxonomyPickerLoader'
import { ITaxonomyObject } from './../interfaces/ITaxonomyObject';
import * as strings from 'OrderGroupWebPartStrings';
import styles from './OrderGroup.module.scss';
import { PrivacySetting } from './PrivacySetting';

export interface IProjectProps {
    cancel: () => void;
    context: WebPartContext;
    updateList: (data: any) => void;
    hideParentDepartment: boolean;
}

const DayPickerStrings: IDatePickerStrings = {
    months: [
        strings.January,
        strings.February,
        strings.March,
        strings.April,
        strings.May,
        strings.June,
        strings.July,
        strings.August,
        strings.September,
        strings.October,
        strings.November,
        strings.December
    ],

    shortMonths: [
        strings.Jan,
        strings.Feb,
        strings.Mar,
        strings.Apr,
        strings.May,
        strings.Jun,
        strings.Jul,
        strings.Aug,
        strings.Sep,
        strings.Oct,
        strings.Nov,
        strings.Dec,
    ],

    days: [
        strings.Sunday,
        strings.Monday,
        strings.Tuesday,
        strings.Wednesday,
        strings.Thursday,
        strings.Friday,
        strings.Saturday
    ],

    shortDays: [
        strings.ShortMonday,
        strings.ShortTuesday,
        strings.ShortWednesday,
        strings.ShortThirsday,
        strings.ShortFriday,
        strings.ShortSaturday,
        strings.ShortSunday,
    ],

    goToToday: strings.goToToday,
    prevMonthAriaLabel: strings.prevMonthAriaLabel,
    nextMonthAriaLabel: strings.nextMonthAriaLabel,
    prevYearAriaLabel: strings.prevYearAriaLabel,
    nextYearAriaLabel: strings.nextYearAriaLabel
};

class Project extends React.Component<IProjectProps, any> {

    constructor(props: any) {
        super(props);
        this.state = {
            projectName: "",
            projectGoal: "",
            projectPurpose: "",
            projectNumber: "",
            customURL: "",
            displayURLField: "none",
            disableProjectNumber: false,
            siteAddress: "",
            parentDepartment: "",
            owningDepartment: null,
            privacySetting: "Closed",
            externalShare: true,
            people: []
        };
    }

    render() {
        return (<div>
            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <label>{strings.NewProject}</label>
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <TextField
                        label={strings.ProjectName}
                        maxLength={255}
                        onChange={(_, newValue) => this._setName(newValue)}
                        required={true}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <Label required={true}>{strings.ProjectLeader}</Label>
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
                        }}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <TextField
                        label={strings.ProjectGoal}
                        maxLength={255}
                        onChange={(_, text: string) => this.setState({ projectGoal: text })}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <TextField
                        label={strings.ProjectPurpose}
                        maxLength={500}
                        onChange={(_, text: string) => this.setState({ projectPurpose: text })}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <TextField
                        label={strings.ProjectNumber}
                        maxLength={20}
                        disabled={this.state.disableProjectNumber}
                        onChange={(_, text: string) => this.setState({ projectNumber: text, siteAddress: text })}
                        required={true}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <Checkbox
                        label={strings.NoProjectNumber}
                        id='checkbox1'
                        onChange={this._onProjectNumberCheckboxChange}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <div style={{ display: this.state.displayURLField }}>
                        <TextField
                            label={strings.PreferedURL}
                            onChange={(_, text: string) => this.setState({ customURL: text, siteAddress: this.convertToSlug(text) })}
                            value={this.state.customURL}
                            required={true}
                        />
                    </div>
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <label>{strings.SiteAddress}: {this.props.context.pageContext.web.absoluteUrl}/prosjekt-{this.state.siteAddress}</label>
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <DatePicker label={strings.StartDate} strings={DayPickerStrings} showWeekNumbers={true} firstWeekOfYear={1} showMonthPickerAsOverlay={true} placeholder='Select a date...'
                        value={this.state.startDate} onSelectDate={newDate => { this.setState({ startDate: newDate }) }}
                    />
                </div>
            </div>

            <div className="ms-Grid-row">
                <div className={styles.positioning}>
                    <DatePicker label={strings.EndDate} strings={DayPickerStrings} showWeekNumbers={true} firstWeekOfYear={1} showMonthPickerAsOverlay={true} placeholder='Select a date...'
                        value={this.state.endDate} onSelectDate={newDate => { this.setState({ endDate: newDate }) }}
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
            

            {/*<div className="ms-Grid-row">
                <div className={styles.positioning + " " + styles.taxonomy}>
                    <TaxonomyPickerLoader
                        context={this.props.context}
                        multi
                        name={strings.Department}
                        onPickerChange={(Name, Option) => { this._onTaxonomyChanged(Name, Option) }}
                        required={true}
                    />
                </div>
                    </div>*/}

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
                        id='checkbox2'
                        defaultChecked={true}
                        onChange={this._onSharingCheckboxChange}
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

    private _onSharingCheckboxChange = (ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {
        if (isChecked == false) {
            this.setState({ externalShare: false });
        }
        else {
            this.setState({ externalShare: true });
        }
    }

    private _onTaxonomyChanged = (Name, Option) => {
        var TaxValue: ITaxonomyObject = { label: "", id: "" };
        TaxValue.label = Option.label;
        TaxValue.id = Option.value;
        this.setState({ owningDepartment: TaxValue })
        console.log(Option);
    }

    private _onProjectNumberCheckboxChange = (ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {
        if (isChecked == false) {
            this.setState({ disableProjectNumber: false });
            this.setState({ displayURLField: "none" });
            this.setState({ customURL: "" });
            this.setState({ siteAddress: "" });
        }
        else {
            let urlFriendlyTitle = '';
            urlFriendlyTitle = this.state.projectName ? this.convertToSlug(this.state.projectName) : '';
            this.setState({
                customURL: this.state.projectName,
                siteAddress: urlFriendlyTitle,
                disableProjectNumber: true,
                displayURLField: "block",
                projectNumber: ""
            });
        }
    }

    private convertToSlug = (Text) =>
    {
        return Text
            .toLowerCase()
            .replace(/[^\w ]+/g,'')
            .replace(/ +/g,'-')
            ;
    }

    private _setName = (title: string) => {
        
        this.setState({
            projectName: title
        });
    }

    private _onFinishClick = () => {
        
        if (this.state.people.length > 0 && this.state.projectName != "") {

            if (!this.props.hideParentDepartment && this.state.parentDepartment == "") {
                alert(strings.IncompleteAlert);
            } else {
                if (this.state.disableProjectNumber && this.state.siteAddress == "") {
                    alert(strings.IncompleteAlert);
                } else if(!this.state.disableProjectNumber && this.state.projectNumber == "") {
                    alert(strings.IncompleteAlert);
                } else {
                    let data = {
                        'ContentTypeId': '0x0100DC8802B232844410BC4B3339D325E8DD',
                        'KDTOOwnerId': this.state.people[0].User.Id,
                        'Title': this.state.projectName,
                        'KDTOProjectGoal': this.state.projectGoal,
                        'KDTOProjectPurpose': this.state.projectPurpose,
                        'KDTOProjectNumber': this.state.projectNumber,
                        'KDTOPreferedUrl': this.state.siteAddress,
                        'KDTOStartDate': this.state.startDate,
                        'KDTOEndDate': this.state.endDate,
                        'KDTOParentDepartment': this.state.parentDepartment,
                        //'KDTOOwningDepartment': { __metadata: { type: "SP.Taxonomy.TaxonomyFieldValue" }, TermGuid: this.state.owningDepartment.id, WssId: -1 },
                        'KDTOSitePrivacy': this.state.privacySetting,
                        'KDTOExternalSharing': false
                    };
                    this.props.updateList(data);
                }
            }

        } else {
            alert(strings.IncompleteAlert);
        }
    }

}

export { Project }