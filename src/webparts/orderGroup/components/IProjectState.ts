import { ITaxonomyObject } from './../interfaces/ITaxonomyObject';
import { SharePointUserPersona } from "../models/PeoplePicker";
import { Key } from "react";

export interface IPRojectState {
    projectName: string;
    projectGoal: string;
    projectPurpose: string;
    projectNumber: string;
    customURL: string;
    displayURLField: string;
    disableProjectNumber: boolean;
    siteAddress: string;
    startDate?: Date;
    endDate?: Date;
    parentDepartment: string;
    owningDepartment?: ITaxonomyObject;
    privacySetting: Key;
    externalShare: boolean;
    createTeam: boolean;
    people?: SharePointUserPersona[];
}
