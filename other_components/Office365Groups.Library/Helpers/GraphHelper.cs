using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Office365Groups.Library
{
	public class GraphHelper
	{
		private static GraphServiceClient graphClient = null;
		// Get an authenticated Microsoft Graph Service client.
		public static GraphServiceClient GetAuthenticatedClient()
		{
			GraphServiceClient graphClient = new GraphServiceClient(
					new DelegateAuthenticationProvider(
							async (requestMessage) =>
							{
								string accessToken = AuthHelper.GetMSGraphAccessToken();
							
								// Append the access token to the request.
								requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
							}));
			return graphClient;
		}

		public static void SignOutClient()
		{
			graphClient = null;
		}


	}
	
	public class GraphUserJson
	{
		public string id { get; set; }
		public string displayName { get; set; }
		public string givenName { get; set; }
		public object jobTitle { get; set; }
		public string mail { get; set; }
		public string mobilePhone { get; set; }
		public object officeLocation { get; set; }
		public string preferredLanguage { get; set; }
		public string surname { get; set; }
		public string userPrincipalName { get; set; }
	}

	public class GraphMembersResult
	{
		public List<GraphUserJson> value { get; set; }
	}
}
