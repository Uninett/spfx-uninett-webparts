import { SharePointUserPersona } from "../components/models/PeoplePicker";
import { ITaxonomyObject } from './interfaces/ITaxonomyObject';

export interface IProjectState {
    errorMessage: string;
    listData: any;
    personObject: any;
    displayNameField: string;
    ownersField: string;
    parentDepartmentField: string;
    ownedDepartmentField?: string;
    projectNumberField: string;
    projectGoalField: string;
    projectPurposeField: string;
    startDateField?: Date;
    endDateField?: Date;
    hideDialog: boolean;
    people?: SharePointUserPersona[];
}