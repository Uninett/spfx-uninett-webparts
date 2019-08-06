using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Office365Groups.Library.ModelValidation;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Graph;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Group = Microsoft.Graph.Group;
using System.Text.RegularExpressions;
using System.Web.Configuration;

namespace Office365Groups.Library
{
	public static class OfficeGroupManager
	{
	    private static readonly string MAIL_PREFIX = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("MAIL_PREFIX")) ? "" : WebConfigurationManager.AppSettings.Get("MAIL_PREFIX");
	    private static readonly string MAIL_SUFFIX = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("MAIL_PREFIX")) ? "" : WebConfigurationManager.AppSettings.Get("MAIL_SUFFIX");
	    private static readonly string TEST_PREFIX = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("TEST_PREFIX")) ? "" : WebConfigurationManager.AppSettings.Get("TEST_PREFIX");
        
        public enum GroupUserEndpoint
		{
			Members,
			Owners 
		};

		public static async Task<List<string>> GetGroupUsers(string groupId, GroupUserEndpoint endpoint, string accessToken)
		{	
			var url = string.Format("https://graph.microsoft.com/v1.0/groups/{0}/" + ((endpoint == GroupUserEndpoint.Members) ? "members": "owners"), groupId);
			HttpClient httpClient = new HttpClient();
			httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
			var response = await httpClient.GetAsync(url);
			var responseData = await response.Content.ReadAsStringAsync();
			GraphMembersResult membersResult = JsonConvert.DeserializeObject<GraphMembersResult>(responseData);

			List <string> members = new List<string>();
			foreach (var member in membersResult.value)
			{
				members.Add(member.userPrincipalName);
			}

			return members;
		}

		public static async Task<Group> GetGroupAsync(string groupId)
		{
			var graphClient = GraphHelper.GetAuthenticatedClient();
			var group = await graphClient.Groups[groupId].Request().GetAsync();
			
			return group;
		}

		public static UnifiedGroupEntity CreateGroup(OfficeGroupInfo groupInfo, string accessToken)
		{
			groupInfo = TransformGroupInfo(groupInfo);
			UnifiedGroupEntity group = new UnifiedGroupEntity();
			try
			{
				group = OfficeGroupUtility.CreateUnifiedGroup(
					groupInfo.DisplayName,
					groupInfo.Description,
					groupInfo.MailNickname,
					accessToken,
					groupInfo.Owners,
					groupInfo.Members,
					groupInfo.IsPrivate
				);
			}
			catch (Microsoft.Graph.ServiceException graphError)
			{
				throw graphError;
			}
			catch (Exception e)
			{
				throw e;
			}

			return group;
		}

		public static UnifiedGroupEntity GetGroup(string groupId, string accessToken)
		{
			UnifiedGroupEntity group = null;
			try
			{
				group = OfficeGroupUtility.GetUnifiedGroup(groupId, accessToken);
			}
			catch (Microsoft.Graph.ServiceException graphError)
			{
				throw graphError;
			}
			catch (Exception e)
			{
				throw e;
			}

			return group;
		}

		public static (Microsoft.SharePoint.Client.ListItem listItem, ClientContext clientContext) GetListItem(int itemId, string listId, string spWebUrl)
		{
			ClientContext clientContext = AuthHelper.GetAppOnlyAuthenticatedContext(spWebUrl);
			var list = clientContext.Web.Lists.GetById(new Guid(listId));
			var listItem = list.GetItemById(itemId);
			clientContext.Load(list);
			clientContext.Load(listItem);
			clientContext.ExecuteQuery();
			return (listItem, clientContext);
		}

        /// <summary>
        /// Extracts {OfficeGroupInfo} object from HttpRequest body and validates
        /// </summary>
        /// <exception cref="ValidationException"></exception>
        /// <param name="req"></param>
        /// <returns></returns>
		public static async Task<OfficeGroupInfo> ExtractGroupInfo(HttpRequestMessage req)
		{
			dynamic body = await req.Content.ReadAsStringAsync();
			OfficeGroupInfo groupInfo = JsonConvert.DeserializeObject<OfficeGroupInfo>(body as string);

            groupInfo.Owners = buildOwnerFromClaims(groupInfo.Owners);


            var result = CustomValidations.Validate(groupInfo);
			if (result.isValid == false)
			{
				var error = result.ValidationExceptions.FirstOrDefault();
				throw new ValidationException(error.Message, error);
			}
            
			return groupInfo;
		}

        public static async Task<MetadataInfo> ExtractMetadata(HttpRequestMessage req)
        {
            dynamic body = await req.Content.ReadAsStringAsync();
            MetadataInfo metadataInfo = JsonConvert.DeserializeObject<MetadataInfo>(body as string);

            var result = CustomValidations.Validate(metadataInfo);
            if (result.isValid == false)
            {
                var error = result.ValidationExceptions.FirstOrDefault();
                throw new ValidationException(error.Message, error);
            }

            return metadataInfo;
        }

        public static string[] buildOwnerFromClaims(string[] ownersClaims)
        {
            string[] owners = ownersClaims;
            
            for(int x = 0; x < owners.Length; x++)
            {
                owners[x] = owners[x].Replace("i:0#.f|membership|", "");
            }

            return owners;
        }

        /// <summary>
        /// Let's use the UnifiedGroupsUtility class from PnP CSOM Core to simplify managed code operations for Office 365 groups
        /// </summary>
        /// <param name="accessToken">Azure AD Access token with Group.ReadWrite.All permission</param>
        public static void ManipulateModernTeamSite(string accessToken)
		{
		    throw new NotImplementedException();
			// Create new modern team site at the url https://[tenant].sharepoint.com/sites/mymodernteamsite
			Stream groupLogoStream = new FileStream("C:\\groupassets\\logo-original.png",
																							FileMode.Open, FileAccess.Read);
			var group = UnifiedGroupsUtility.CreateUnifiedGroup("displayName", "description",
															"mymodernteamsite", accessToken, groupLogo: groupLogoStream);
			
			// We received a group entity containing information about the group
			string url = group.SiteUrl;
			string groupId = group.GroupId;

			// Get group based on groupID
			var group2 = UnifiedGroupsUtility.GetUnifiedGroup(groupId, accessToken);
			// Get SharePoint site URL from group id
			var siteUrl = UnifiedGroupsUtility.GetUnifiedGroupSiteUrl(groupId, accessToken);

			// Get all groups in the tenant
			List<UnifiedGroupEntity> groups = UnifiedGroupsUtility.ListUnifiedGroups(accessToken);

			// Update description and group logo programatically
			groupLogoStream = new FileStream("C:\\groupassets\\logo-new.png", FileMode.Open, FileAccess.Read);
			UnifiedGroupsUtility.UpdateUnifiedGroup(groupId, accessToken, description: "Updated description",
																							groupLogo: groupLogoStream);

			// Delete group programatically
			UnifiedGroupsUtility.DeleteUnifiedGroup(groupId, accessToken);
		}

        /// <summary>
        /// Updates {OfficeGroupInfo} object with correct MailNickname and prefix.
        /// </summary>
        /// <param name="groupInfo"></param>
        /// <returns></returns>
		private static OfficeGroupInfo TransformGroupInfo(OfficeGroupInfo groupInfo)
		{
            // build mail alias from MailNickname
               
            // Add områdetype prefix her!

			var mailNickname = ConvertScandiChars(groupInfo.MailNickname);
		    if (!string.IsNullOrEmpty(MAIL_PREFIX)) mailNickname = ConvertScandiChars(MAIL_PREFIX) + mailNickname;
		    if (!string.IsNullOrEmpty(MAIL_SUFFIX)) mailNickname = mailNickname + ConvertScandiChars(MAIL_SUFFIX);

            // Test prefix
		    if (!string.IsNullOrEmpty(TEST_PREFIX)) mailNickname = ConvertScandiChars(TEST_PREFIX) + mailNickname;

            groupInfo.MailNickname = mailNickname;
            groupInfo.Description = string.IsNullOrEmpty(groupInfo.Description) ? " " : groupInfo.Description;
            return groupInfo;
		}

	    private static string ConvertScandiChars(string scandis = "")
	    {
            // allowed mail chars
	        Regex rgx = new Regex("[^a-zA-Z0-9-_]");

            // string scandis = "æ, ø, å, Æ, Ø, Å ";
            scandis = Regex.Replace(scandis, "[æøåÆØÅ]", (m) => 
	            (m.Value == "æ") ? "ae" :
	            (m.Value == "ø") ? "o" :
	            (m.Value == "å") ? "a" :
	            (m.Value == "Æ") ? "AE" :
	            (m.Value == "Ø") ? "O" :
	            (m.Value == "Å") ? "A" : m.Value);
            
            return rgx.Replace(scandis, "");
	    }

    }
}
