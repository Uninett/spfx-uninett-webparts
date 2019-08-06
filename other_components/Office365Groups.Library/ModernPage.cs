using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office365Groups.Library
{
	public class ModernPage
	{
		private string siteUrl;
		private string pageName;
		private string accessToken;
		private ClientContext clientContext { get; set; }
		private ClientSidePage page { get; set; }

		public void ModenPage() {
			AuthenticationManager am = new AuthenticationManager();
			accessToken = AuthHelper.GetMSGraphAccessToken();
			clientContext = am.GetAzureADAccessTokenAuthenticatedContext(siteUrl, this.accessToken);
		}

		public ClientSidePage ModenPage(string siteUrl, string pageName)
		{
			AuthenticationManager am = new AuthenticationManager();
			accessToken = AuthHelper.GetMSGraphAccessToken();
			clientContext = am.GetAzureADAccessTokenAuthenticatedContext(siteUrl, this.accessToken);
			page = new ClientSidePage(clientContext);

			return page;
		}
		
		private void createPage(string pageName, ClientContext siteContext = null)
		{
			siteContext = siteContext == null ? clientContext : siteContext ;
			page = new ClientSidePage(siteContext);

			ClientSideText txt1 = new ClientSideText() { Text = "COB test" };
			page.AddControl(txt1, 0);

			// page will be created if it doesn't exist, otherwise overwritten if it does..
			page.Save(pageName);
		}

		private void modifyExistingPage(string pageName, ClientContext siteContext)
		{
			siteContext = siteContext == null ? clientContext : siteContext;

			// load exising page - will return null if no page found with this name..
			ClientSidePage page = ClientSidePage.Load(siteContext, pageName);

			ClientSideText txt1 = new ClientSideText() { Text = "COB test" };
			page.AddControl(txt1, 0);

			page.Save(pageName);
		}

		private ClientSidePage getPage(string pageName, ClientContext siteContext)
		{
			siteContext = siteContext == null ? clientContext : siteContext;

			// load exising page - will return null if no page found with this name..
			ClientSidePage page = ClientSidePage.Load(siteContext, pageName);
			return page;
		}

	}
}
