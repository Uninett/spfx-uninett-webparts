import { SharePointUserPersona } from "../models/PeoplePicker";
import { Key } from "react";

export interface IActivityState {
    activityName: string;
    activityDescription: string;
    parentDepartment: string;
    privacySetting: Key;
    externalShare: boolean;
    people?: SharePointUserPersona[];
}