import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISiteCatalogProps {
  listRows: string;
  showNewButton: boolean;
  context: WebPartContext;
  siteTypes: string;
  hideParentDepartment: boolean;
  searchType: SearchType;
}
