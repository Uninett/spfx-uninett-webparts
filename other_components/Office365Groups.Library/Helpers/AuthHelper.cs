using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Configuration;
using ClientCredential = Microsoft.Identity.Client.ClientCredential;
using AuthenticationResult = Microsoft.Identity.Client.AuthenticationResult;
using Microsoft.SharePoint.Client;
using System.Security;

namespace Office365Groups.Library
{
	
	public static class AuthHelper
	{

		#region AuthenticationProperties
		private static string RedirectUri = "urn:ietf:wg:oauth:2.0:oob";
		private static Uri AADLogin = new Uri("https://login.microsoftonline.com/");
		private static string[] AADDefaultScope = { "https://graph.microsoft.com/.default" };

		// Client secrets for SharePoint registered App
		public static string SPClientId = WebConfigurationManager.AppSettings.Get("SPClientId");
		public static string SPClientSecret = WebConfigurationManager.AppSettings.Get("SPClientSecret");

		// Client secrets for apps.dev.microsoft.com (AAD v2) registered APP.
		private static string AADClientId = WebConfigurationManager.AppSettings.Get("AADClientId");
		private static string AADClientSecret = WebConfigurationManager.AppSettings.Get("AADClientSecret");
		private static string AADDomain = WebConfigurationManager.AppSettings.Get("AADDomain");

		// SharePoint Authenticated User
		public static string SPUserId = WebConfigurationManager.AppSettings.Get("SPUserId");
		public static string SPUserPassword = WebConfigurationManager.AppSettings.Get("SPUserPassword");

		public static string SPGroupsSiteUrl = WebConfigurationManager.AppSettings.Get("SPGroupsSiteUrl");
        public static string TenantAdminUrl = WebConfigurationManager.AppSettings.Get("TenantAdminUrl");

	    private static string CompanyName = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CompanyName")) ? AADDomain : WebConfigurationManager.AppSettings.Get("CompanyName");

	    public static string TRAFFIC_DECORATION =
	        $"NONISV|{CompanyName}|OfficeGroupsProvisioning/1.0";
        #endregion

        #region PublicMembers
        public static ClientContext GetAppOnlyAuthenticatedContext(string targetPricipleName)
		{
			OfficeDevPnP.Core.AuthenticationManager am = new OfficeDevPnP.Core.AuthenticationManager();
			ClientContext clientContext = am.GetAppOnlyAuthenticatedContext(targetPricipleName, SPClientId, SPClientSecret);

			return clientContext;
		}

		public static ClientContext GetSharePointClientContext(string siteUrl = "")
		{
			siteUrl = string.IsNullOrEmpty(siteUrl) ? SPGroupsSiteUrl : siteUrl ;
			ClientContext ctx = new ClientContext(siteUrl);
			var passWord = new SecureString();
			foreach (char c in SPUserPassword.ToCharArray())
			{
				passWord.AppendChar(c);
			}

			ctx.Credentials = new SharePointOnlineCredentials(SPUserId, passWord);
		    ctx.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
		    {
		        e.WebRequestExecutor.WebRequest.UserAgent = TRAFFIC_DECORATION;
		    };

            return ctx;
		}

        public static ClientContext GetTenantAdminClientContext()
        {
            string tenantAdminUrl = TenantAdminUrl;
            ClientContext ctx = new ClientContext(tenantAdminUrl);
            var passWord = new SecureString();
            foreach (char c in SPUserPassword.ToCharArray())
            {
                passWord.AppendChar(c);
            }

            ctx.AuthenticationMode = ClientAuthenticationMode.Default;
            ctx.Credentials = new SharePointOnlineCredentials(SPUserId, passWord);
            ctx.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
            {
                e.WebRequestExecutor.WebRequest.UserAgent = TRAFFIC_DECORATION;
            };

            return ctx;
        }

		public static string GetMSGraphAccessToken()
		{
			var appCredentials = new ClientCredential(AADClientSecret);
			var authority = new Uri(AADLogin, AADDomain).AbsoluteUri;
			var clientApplication = new ConfidentialClientApplication(AADClientId, authority, RedirectUri, appCredentials, null, null);
			AuthenticationResult authenticationResult = clientApplication.AcquireTokenForClientAsync(AADDefaultScope).GetAwaiter().GetResult();

			string accessToken = authenticationResult.AccessToken;

			return accessToken;
		}

		#endregion

	}
}
