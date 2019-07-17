import { SharePointUserPersona } from "../components/models/PeoplePicker";

export interface IActivityState {
    errorMessage: string;
    listData: any;
    personObject: any;
    displayNameField: string;
    descriptionField: string;
    ownersField: string;
    parentDepartmentField: string;
    hideDialog: boolean;
    people?: SharePointUserPersona[];
}