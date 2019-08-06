import { SharePointUserPersona } from "../components/models/PeoplePicker";

export interface IGroupState {
    errorMessage: string;
    listData: any;
    personObject: any;
    displayNameField: string;
    ownersField: string;
    hideDialog: boolean;
    people?: SharePointUserPersona[];
}