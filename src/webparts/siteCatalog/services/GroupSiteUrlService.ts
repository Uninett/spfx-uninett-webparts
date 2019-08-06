import { IGroupSiteUrl } from '../components/interfaces/IGroupSiteUrl';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, HttpClientResponse, IGraphHttpClientOptions, GraphHttpClient, AadHttpClient } from '@microsoft/sp-http';
 
interface IGroupSiteUrlService {
  getGroupSiteUrl: Promise<IGroupSiteUrl>;
  webAbsoluteUrl: string;
  context: WebPartContext;
  groupId: string;
}
 
export class GroupSiteUrlService {
 
  private context: WebPartContext;
  private groupId: string;
 
  constructor(_context, _groupId) {
    this.context = _context;
    this.groupId = _groupId;
  }
 
  public getGroupSiteUrl(): Promise<IGroupSiteUrl> {

    return this.context.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient): Promise<IGroupSiteUrl> => {
        return client
          .get("https://graph.microsoft.com/v1.0/groups/" + this.groupId + "/sites/root", AadHttpClient.configurations.v1)
          .then((response: HttpClientResponse) => {
            return response.json();
          });
    });

    /*
    return this.context.graphHttpClient.get("v1.0/groups/" + this.groupId + "/sites/root", GraphHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        return response.json();
    });
    */

  }
}