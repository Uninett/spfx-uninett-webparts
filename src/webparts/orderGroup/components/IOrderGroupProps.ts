import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SiteType } from '../interfaces/SiteType';

export interface IOrderGroupProps {
  description: string;
  context: WebPartContext;
  departmentSiteTypeName: SiteType;
  hideParentDepartment: boolean;
}
