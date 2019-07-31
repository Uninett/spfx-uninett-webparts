import { SharePointUserPersona } from "../models/PeoplePicker";
import { Key } from "react";

export interface IDepartmentState {
    departmentName: string;
    departmentDescription: string;
    // shortName: string;
    privacySetting: Key;
    externalShare: boolean;
    people?: SharePointUserPersona[];
}