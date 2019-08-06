import { IFavoriteSites } from '../interfaces/IFavoriteSites';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, HttpClientResponse, IGraphHttpClientOptions } from '@microsoft/sp-http';
 
interface IFavoriteSitesService {
  getFavoriteSites: Promise<IFavoriteSites>;
  webAbsoluteUrl: string;
  context: WebPartContext;
}
 
export class FavoriteSitesService {
 
  private context: WebPartContext;
 
  constructor(_context) {
    this.context = _context;
  }
 
  public getFavoriteSites(): Promise<IFavoriteSites> {

      return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_vti_bin/homeapi.ashx/sites/followed`, SPHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        return response.json();
      });
  }
}