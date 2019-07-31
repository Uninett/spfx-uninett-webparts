import * as React from "react";
import { DefaultButton, CompoundButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import * as strings from 'OrderGroupWebPartStrings';
import styles from './OrderGroup.module.scss';
import ActivityImage from './assets/ActivityImage';
import DepartmentImage from './assets/DepartmentImage';
import ProjectImage from './assets/ProjectImage';
import { SiteType } from '../interfaces/SiteType';

export interface IGroupTypeChoiceProps {
    onChange: (siteType: string) => void;
    departmentSiteTypeName: SiteType;
}

class GroupTypeChoice extends React.Component<IGroupTypeChoiceProps, any> {
    public render() {
        let deparmentName = this.props.departmentSiteTypeName === 0 ? strings.Department : strings.Section;

        return (
            <div>
                <div className="ms-Grid-row">
                    <div className={styles.positioning}>
                        <span>{strings.ChooseType}</span>
                    </div>
                </div>

                <div className="ms-Grid-row">
                    <div className={styles.positioning}>
                        <CompoundButton onClick={() => this.props.onChange("Activity")}>
                            <div className="ms-Image">
                                <p>{strings.Activity}</p>
                                <ActivityImage/>
                            </div>
                        </CompoundButton>
                    </div>

                    <div className={styles.positioning}>
                        <CompoundButton onClick={() => this.props.onChange("Department")}>
                            <div className="ms-Image">
                                <p>{deparmentName}</p>
                                <DepartmentImage/>
                            </div>
                        </CompoundButton>
                    </div>

                    <div className={styles.positioning}>
                        <CompoundButton onClick={() => this.props.onChange("Project")}>
                            <div className="ms-Image">
                                <p>{strings.Project}</p>
                                <ProjectImage/>
                            </div>
                        </CompoundButton>
                    </div>
                </div>
            </div>

        );
    }
}


export { GroupTypeChoice };