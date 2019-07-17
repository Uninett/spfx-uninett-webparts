import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";

export interface ISiteMetadataProps {
  description: string;
  orderSiteURL: string;
  editName: boolean;
  editDescription: boolean;
  editProjectGoal: boolean;
  editProjectPurpose: boolean;
  editOwner: boolean;
  editShortName: boolean;
  editParentDepartment: boolean;
  editOwningDepartment: boolean;
  editProjectNumber: boolean;
  editStartDate: boolean;
  editEndDate: boolean;
  context: WebPartContext;
  displayMode: DisplayMode;
}
