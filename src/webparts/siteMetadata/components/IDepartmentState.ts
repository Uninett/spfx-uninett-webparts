import { SharePointUserPersona } from "../components/models/PeoplePicker";

export interface IDepartmentState {
    errorMessage: string;
    listData: any;
    personObject: any;
    displayNameField: string;
    descriptionField: string;
    ownersField: string;
    parentDepartmentField: string;
    shortNameField: string;
    hideDialog: boolean;
    people?: SharePointUserPersona[];
}