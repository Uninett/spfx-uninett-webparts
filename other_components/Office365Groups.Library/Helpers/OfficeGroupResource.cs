using System;
using System.Net;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Office365Groups.Library.Models;
using J = Newtonsoft.Json.JsonPropertyAttribute;

namespace Office365Groups.Library.Helpers
{
    public class OfficeGroupResource
    {
        [J("alias")] public string Alias { get; set; }
        [J("allowToAddGuests")] public bool AllowToAddGuests { get; set; }
        [J("calendarUrl")] public string CalendarUrl { get; set; }
        [J("classification")] public object Classification { get; set; }
        [J("description")] public string Description { get; set; }
        [J("displayName")] public string DisplayName { get; set; }
        [J("documentsUrl")] public string DocumentsUrl { get; set; }
        [J("editGroupUrl")] public string EditGroupUrl { get; set; }
        [J("id")] public string Id { get; set; }
        [J("inboxUrl")] public string InboxUrl { get; set; }
        [J("isDynamic")] public bool IsDynamic { get; set; }
        [J("isPublic")] public bool IsPublic { get; set; }
        [J("mail")] public string Mail { get; set; }
        [J("notebookUrl")] public object NotebookUrl { get; set; }
        [J("peopleUrl")] public string PeopleUrl { get; set; }
        [J("pictureUrl")] public string PictureUrl { get; set; }
        [J("principalName")] public object PrincipalName { get; set; }
        [J("siteUrl")] public string SiteUrl { get; set; }
        [J("owner")] public string Owner { get; set; }

        public static OfficeGroupResource Get(ClientContext context, string groupId)
        {
            context.Load(context.Web, w => w.Url);
            context.ExecuteQuery();

            using (var handler = new HttpClientHandler())
            {
                // Set permission setup accordingly for the call
                handler.Credentials = context.Credentials;
                handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));

                using (var httpClient = new HttpClient(handler))
                {
                    string apiUrl =
                        "_api/SP.Directory.DirectorySession/Group('" + groupId + "')?$select=PrincipalName,Id,DisplayName,Alias,Description,InboxUrl,CalendarUrl,DocumentsUrl,SiteUrl,EditGroupUrl,PictureUrl,PeopleUrl,NotebookUrl,Mail,IsPublic,CreationTime,Classification,allowToAddGuests,isDynamic";
                    
                    string requestUrl = $"{context.Web.Url}/{apiUrl}";
                    
                    httpClient.DefaultRequestHeaders.Add("accept", "application/json;odata=nometadata");
                    var response = httpClient.GetAsync(requestUrl).GetAwaiter().GetResult();

                    if (response.IsSuccessStatusCode)
                    {
                        var responseString = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                        return JsonConvert.DeserializeObject<OfficeGroupResource>(responseString);
                    }

                    // Something went wrong...
                    throw new Exception(response.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                }
            }
        }
        
    }
}
