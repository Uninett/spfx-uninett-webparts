//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Facebook;
//using System.Web;
//using System.Net.Http;

//namespace NRKGroups.Library
//{
//	public class WorkplaceClient
//	{

//		private static readonly HttpClient httpClient = new HttpClient();
//		// Constants
//		public const string GRAPH_URL_PREFIX = "https://graph.facebook.com/";
//		public const string FIELDS_CONJ = "?fields=";
//		public const string GROUPS_SUFFIX = "/groups";
//		public const string GROUP_FIELDS = "id,name,members,privacy,description,updated_time";
//		public const string MEMBERS_SUFFIX = "/members";
//		public const string MEMBER_FIELDS = "email,id,administrator";
//		public const string JSON_KEY_DATA = "data";
//		public const string JSON_KEY_PAGING = "paging";
//		public const string JSON_KEY_NEXT = "next";
//		public const string JSON_KEY_EMAIL = "email";

//		private FacebookClient fb;

//		public WorkplaceClient() {
//			var accessToken = "";
//			fb = new FacebookClient(accessToken);
//		}

//		public object CreateNewGroup(string communityId, string name, string description, string privacy, string admin)
//		{
//			var endpoint = GRAPH_URL_PREFIX + communityId + GROUPS_SUFFIX;
//			object newGroup = null;
//			try
//			{
//				var url = endpoint + string.Format("?name={0}&description={1}&privacy={2}&owner={3}", HttpUtility.UrlEncode(name), HttpUtility.UrlEncode(description), privacy, admin);
//				var accessToken = "";
//				httpClient.DefaultRequestHeaders.Add("Authentication", "Bearer " + accessToken);
//				var response = httpClient.PostAsync(url, null).Result;
//				var data = response.Content.ReadAsStringAsync();
//				return data;
//			}
//			catch (Exception e)
//			{
//				var ex = e;
//				throw ex;
//			}
//			return newGroup;
//		}

//		public object GetAllGroups(string communityId)
//		{
//			var endpoint = GRAPH_URL_PREFIX + communityId + GROUPS_SUFFIX + FIELDS_CONJ + GROUP_FIELDS;
//			return fb.Get(endpoint);
//		}
		
//	}

//	public class WorkplaceGroupInfo
//	{
//		public string name { get; set; }
//		public string description { get; set; }
//		public string privacy { get; set; }
//		public string admin { get; set; }
//	}
//}
