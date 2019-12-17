using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Office365Groups.Library;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Collections.Generic;
using OfficeDevPnP.Core.Entities;
using Microsoft.SharePoint.Client;
using System.Linq;
using Office365Groups.Library.Helpers;
using Group = Microsoft.Graph.Group;
using Team = Microsoft.Graph.Team;

namespace Rederi.Functions
{
    public static class CreateOfficeGroup
    {
        public static TraceWriter FnLog { get; set; }

        [FunctionName("CreateOfficeGroup")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]HttpRequestMessage req, ExecutionContext context, TraceWriter log)
        {
            FnLog = log;
            OfficeGroupUtility.Logger = FnLog;
            UnifiedGroupEntity currentGroup = null;
            var itemId = "";
            // extract listitem properties from "request"
            var groupReq = await OfficeGroupManager.ExtractGroupInfo(req);
            var owner = groupReq.Owners[0];
            groupReq.Validate();
            itemId = groupReq.DisplayName;
            OfficeGroupResource createdGroup;
            HttpStatusCode createGroupStatus;

            try
            {
                var accessToken = AuthHelper.GetMSGraphAccessToken();
                if (string.IsNullOrEmpty(groupReq.GroupId))
                {
                    var groupInfo = groupReq;
                    FnLog.Info($"Starting provisioning  >>>");

                    FnLog.Info($"Creating new group with name: {groupInfo.DisplayName}  >>>");
                    currentGroup = OfficeGroupManager.CreateGroup(groupInfo, accessToken);

                    FnLog.Info($"Group created with Id: {currentGroup.GroupId} and Url: {currentGroup.SiteUrl}  >>>");

                    // getting siteurl may take time.
                    groupInfo.SiteUrl = EnsureGroupSiteUrl(currentGroup, accessToken);

                    groupInfo.GroupId = currentGroup.GroupId;

                    if (groupInfo.CreateTeam)
                    {
                        FnLog.Info($"Creating team >>>");
                        CreateTeamFromGroup(groupInfo.GroupId);
          }
                    
                }
                else
                {
                    if (string.IsNullOrEmpty(groupReq.GroupId))
                        throw new KeyNotFoundException("Unexpected provsioned group without GroupId");

                    CreateOfficeGroup.FnLog.Info($"OrderId: {itemId} - Updating a group is no longer supported >>>");


                    return req.CreateResponse(HttpStatusCode.OK);

                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);

                return ErrorHandler.CreateExceptionResponse(req, ex);
            }

            CreateOfficeGroup.FnLog.Info($"V");
            CreateOfficeGroup.FnLog.Info($"V");
            CreateOfficeGroup.FnLog.Info($"V");
            using (var ctx = AuthHelper.GetSharePointClientContext())
            {
                try
                {
                    var groupResource = OfficeGroupResource.Get(ctx, currentGroup.GroupId);
                    createdGroup = groupResource;
                    createdGroup.Owner = owner;
                    createGroupStatus = HttpStatusCode.OK;


                    // Set external sharing on group
                    var accessToken = AuthHelper.GetMSGraphAccessToken();
                    SetSharingCapabilities(accessToken, createdGroup.Id, groupReq.ExternalSharing);


                }
                catch (Exception e)
                {

                    return ErrorHandler.CreateExceptionResponse(req, e);
                }
            }

            try
            {
                return req.CreateResponse(HttpStatusCode.OK, createdGroup);
            }
            catch (Exception e)
            {
                return ErrorHandler.CreateExceptionResponse(req, e);
            }


        }


        public static void SetSharingCapabilities(string accessToken, string groupId, bool sharing)
        {

            // Disable/Enable sharing on group
            OfficeGroupUtility.AllowToAddGuests(groupId, sharing, accessToken);
            CreateOfficeGroup.FnLog.Info($"External sharing set to: {sharing}");

        }

    public static async void CreateTeamFromGroup(string groupId)
    {
      var graphClient = GraphHelper.GetAuthenticatedClient();

      var team = new Team
      {
        AdditionalData = new Dictionary<string, object>()
            {
              {"group@odata.bind","https://graph.microsoft.com/v1.0/groups('" + groupId + "')"},
              {"template@odata.bind","https://graph.microsoft.com/beta/teamsTemplates('standard')"}
            }
      };

      await graphClient.Teams
        .Request()
        .AddAsync(team);
    }

    /// <summary>
    /// Returns the URL of the Modern SharePoint Site backing an Office 365 Group (i.e. Unified Group)
    /// </summary>
    /// <param name="currentGroup">The Office 365 Group</param>
    /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
    /// <returns>The URL of the modern site backing the Office 365 Group</returns>
    private static string EnsureGroupSiteUrl(UnifiedGroupEntity currentGroup, string accessToken)
        {
            if (string.IsNullOrEmpty(currentGroup.SiteUrl))
            {
                CreateOfficeGroup.FnLog.Warning($"Cannot find siteUrl for group: {currentGroup.GroupId} ");
                CreateOfficeGroup.FnLog.Warning($"2nd trial to find SiteUrl for group: {currentGroup.GroupId} ");
                try
                {
                    return OfficeGroupUtility.GetUnifiedGroupSiteUrl(currentGroup.GroupId, accessToken);
                }
                catch (Exception e)
                {
                    CreateOfficeGroup.FnLog.Error($"Failed to get SiteUrl for group: {currentGroup.GroupId} ", e);
                    return null;
                }
            }
            else
            {
                return currentGroup.SiteUrl;
            }
        }

        private static (ListItem listItem, ClientContext clientContext) SetPermissions(
            (ListItem listItem, ClientContext clientContext) itemData)
        {
            var listItem = itemData.listItem;
            var clientContext = itemData.clientContext;
            var listItemId = listItem.Id;

            try
            {
                // Break inheritance and remove current assignments
                listItem.BreakRoleInheritance(false, true);
                try
                {
                    // TODO: Remove the following - it threw @cannot find principal with ID 3 error
                    for (int i = 0; i < listItem.RoleAssignments.Count; i++)
                    {
                        listItem.RoleAssignments[i].RoleDefinitionBindings.RemoveAll();
                    }
                }
                catch
                {
                    // known cannot find principal with ID 3 error... continue
                }

                listItem.SystemUpdate();
                clientContext.Load(listItem, i => i.RoleAssignments);
                clientContext.ExecuteQuery();
                CreateOfficeGroup.FnLog.Info($"OrderId: {listItemId} - Removed listItem permissions.");
            }
            catch (Exception e)
            {
                CreateOfficeGroup.FnLog.Warning($"OrderId: {listItemId} - Error clearing permissions", e.Message);

                // continue -- no need to throw error
            }

            try
            {
                //Microsoft.SharePoint.Client.Group approvalGroup = null;
                //try
                //{
                //	approvalGroup = clientContext.Web.SiteGroups.GetByName("Godkjennere");
                //	clientContext.Load(approvalGroup);
                //	clientContext.ExecuteQuery();
                //}
                //catch (Exception e)
                //{
                //	CreateGroup.FnLog.Error($"OrderId: {listItemId} - Error getting Godkjennere group.", e);
                //}

                // Add listItem author
                FieldUserValue authorsValue = (FieldUserValue) listItem["Author"];
                Microsoft.SharePoint.Client.User principal = clientContext.Web.GetUserById(authorsValue.LookupId);
                SetListItemRole(listItem, clientContext, principal, RoleType.Administrator);

                // Add Owners
                FieldUserValue[] ownersValue = listItem["NrkGroupOwner"] as FieldUserValue[];
                foreach (FieldUserValue ownerValue in ownersValue)
                {
                    Microsoft.SharePoint.Client.User ownerUser = clientContext.Web.GetUserById(ownerValue.LookupId);
                    SetListItemRole(listItem, clientContext, ownerUser, RoleType.Contributor);
                }

                // Add Members
                FieldUserValue[] membersValue = listItem["NrkMembers"] as FieldUserValue[];
                foreach (FieldUserValue memberValue in membersValue)
                {
                    Microsoft.SharePoint.Client.User memberUser = clientContext.Web.GetUserById(memberValue.LookupId);
                    SetListItemRole(listItem, clientContext, memberUser, RoleType.Reader);
                }

                // Add Visitors
                if (listItem["NrkPrivacy"]?.ToString() == "Offentlig")
                {
                    SetListItemRole(listItem, clientContext, clientContext.Web.AssociatedVisitorGroup, RoleType.Reader);
                    SetListItemRole(listItem, clientContext, clientContext.Web.AssociatedMemberGroup, RoleType.Reader);
                }

                // Add godkjennere to structured groups
                if (listItem["NrkGroupType"]?.ToString() == "Strukturert")
                {
                    // check if godkjennere sharepoint group exist
                    Microsoft.SharePoint.Client.Group godkjennere =
                        clientContext.Web.SiteGroups.GetByName("Godkjennere");

                    if (godkjennere != null)
                    {
                        // if it does, set role on item
                        SetListItemRole(listItem, clientContext, godkjennere, RoleType.Contributor);
                    }
                }
            }
            catch (Exception e)
            {
                CreateOfficeGroup.FnLog.Error($"OrderId: {listItemId} - Error adding unique permissions.", e);
                CreateOfficeGroup.FnLog.Warning($"OrderId: {listItemId} - Resetting unique permission.");
                listItem.ResetRoleInheritance();
            }

            try
            {
                // this will not trigger Workflows
                listItem.SystemUpdate();
                clientContext.Load(listItem);
                clientContext.ExecuteQuery();
                CreateOfficeGroup.FnLog.Info($"OrderId: {listItemId} - Removed listItem permissions.");
            }
            catch (Exception e)
            {
                CreateOfficeGroup.FnLog.Error($"OrderId: {listItemId} - Error adding unique permissions.", e);
                CreateOfficeGroup.FnLog.Warning($"OrderId: {listItemId} - Resetting unique permission.");
                listItem.ResetRoleInheritance();
                listItem.SystemUpdate();
                clientContext.Load(listItem);
                clientContext.ExecuteQuery();
            }

            itemData.listItem = listItem;
            itemData.clientContext = clientContext;

            return itemData;
        }


        /// <summary>
        /// Updates order ListItem of an Office 365 Group
        /// </summary>
        /// <param name="listItem">The SharePoint ListItem that holdes the order information</param>
        /// <param name="clientContext">SharePoint ClientContext related to the ListItem</param>
        /// <param name="officeGroup">OfficeGroupInfo object which holdes the group information</param>
        /// <returns></returns>
        private static void UpdateGroupListItem(ListItem listItem, ClientContext clientContext,
            OfficeGroupInfo officeGroup)
        {
            // Update list back.
            var listItemId = listItem.Id;
            if (string.IsNullOrEmpty(listItem["NrkGroupId"] as string))
                listItem["NrkGroupId"] = officeGroup?.GroupId;

            if (!string.IsNullOrEmpty(officeGroup?.SiteUrl))
                listItem["NrkGroupUrl"] = officeGroup?.SiteUrl;

            listItem["NrkProvisioned"] = "Ja";

            listItem.SystemUpdate();
            clientContext.Load(listItem, i => i.RoleAssignments);
            clientContext.ExecuteQuery();
            CreateOfficeGroup.FnLog.Info($"OrderId: {listItemId} - Updated GroupId and Url");

            /*try
            {
                // Break inheritance and remove current assignments
                listItem.BreakRoleInheritance(false, true);
                try
                {
                    // TODO: Remove the following - it threw @cannot find principal with ID 3 error
                    for (int i = 0; i < listItem.RoleAssignments.Count; i++)
                    {
                        listItem.RoleAssignments[i].RoleDefinitionBindings.RemoveAll();
                    }
                } catch
                {
                    // known cannot find principal with ID 3 error... continue
                }

                listItem.SystemUpdate();
                clientContext.Load(listItem, i => i.RoleAssignments);
                clientContext.ExecuteQuery();
                CreateGroup.FnLog.Info($"OrderId: {listItemId} - Removed listItem permissions.");

            }
            catch (Exception e)
            {
                CreateGroup.FnLog.Warning($"OrderId: {listItemId} - Error clearing permissions", e.Message);

                // continue -- no need to throw error
            }


            try
            {
                //Microsoft.SharePoint.Client.Group approvalGroup = null;
                //try
                //{
                //	approvalGroup = clientContext.Web.SiteGroups.GetByName("Godkjennere");
                //	clientContext.Load(approvalGroup);
                //	clientContext.ExecuteQuery();
                //}
                //catch (Exception e)
                //{
                //	CreateGroup.FnLog.Error($"OrderId: {listItemId} - Error getting Godkjennere group.", e);
                //}

                // Add listItem author
                FieldUserValue authorsValue = (FieldUserValue)listItem["Author"];
                User principal = clientContext.Web.GetUserById(authorsValue.LookupId);
                SetListItemRole(listItem, clientContext, principal, RoleType.Administrator);

                // Add Owners
                FieldUserValue[] ownersValue = listItem["NrkGroupOwner"] as FieldUserValue[];
                foreach (FieldUserValue ownerValue in ownersValue)
                {
                    User ownerUser = clientContext.Web.GetUserById(ownerValue.LookupId);
                    SetListItemRole(listItem, clientContext, ownerUser, RoleType.Contributor);
                }

                // Add Members
                FieldUserValue[] membersValue = listItem["NrkMembers"] as FieldUserValue[];
                foreach (FieldUserValue memberValue in membersValue)
                {
                    User memberUser = clientContext.Web.GetUserById(memberValue.LookupId);
                    SetListItemRole(listItem, clientContext, memberUser, RoleType.Reader);
                }

                // Add Visitors
                if (officeGroup.IsPrivate == false)
                {
                    SetListItemRole(listItem, clientContext, clientContext.Web.AssociatedVisitorGroup, RoleType.Reader);
                }

                                // Add godkjennere to structured groups
                                if (officeGroup.GroupType == GroupType.Structured)
                                {
                                        // check if godkjennere sharepoint group exist
                                        Microsoft.SharePoint.Client.Group godkjennere = clientContext.Web.SiteGroups.GetByName("Godkjennere");

                                        if (godkjennere != null)
                                        {
                                                // if it does, set role on item
                                                SetListItemRole(listItem, clientContext, godkjennere, RoleType.Contributor);
                                        }


                                }

            }
            catch (Exception e)
            {
                CreateGroup.FnLog.Error($"OrderId: {listItemId} - Error adding unique permissions.", e);
                CreateGroup.FnLog.Warning($"OrderId: {listItemId} - Resetting unique permission.");
                listItem.ResetRoleInheritance();
            }

            try
            {
                // this will not trigger Workflows
                listItem.SystemUpdate();
                clientContext.Load(listItem);
                clientContext.ExecuteQuery();
                CreateGroup.FnLog.Info($"OrderId: {listItemId} - Removed listItem permissions.");
            }
            catch (Exception e)
            {
                CreateGroup.FnLog.Error($"OrderId: {listItemId} - Error adding unique permissions.", e);
                CreateGroup.FnLog.Warning($"OrderId: {listItemId} - Resetting unique permission.");
                listItem.ResetRoleInheritance();
                listItem.SystemUpdate();
                clientContext.Load(listItem);
                clientContext.ExecuteQuery();

            }*/
        }

        private static void SetListItemRole(ListItem listItem, ClientContext clientContext, Principal principal,
            RoleType role)
        {
            RoleDefinitionBindingCollection collRoleDefinitionBinding =
                new RoleDefinitionBindingCollection(clientContext)
                {
                    clientContext.Web.RoleDefinitions.GetByType(role) //Set permission type
                };
            listItem.RoleAssignments.Add(principal, collRoleDefinitionBinding);
        }

        private static UnifiedGroupEntity UpdateGroup(OfficeGroupInfo oldGroupInfo, OfficeGroupInfo newGroupInfo,
            Group oldGroup)
        {
            // No need to update
            // return oldGroup.
            if (
                newGroupInfo.DisplayName == oldGroupInfo.DisplayName &&
                newGroupInfo.Description == oldGroupInfo.Description &&
                newGroupInfo.IsPrivate == oldGroupInfo.IsPrivate &&
                newGroupInfo.Owners.SequenceEqual(oldGroupInfo.Owners) &&
                newGroupInfo.Members.SequenceEqual(oldGroupInfo.Members)
            )
            {
                return new UnifiedGroupEntity
                {
                    GroupId = oldGroupInfo.GroupId,
                    DisplayName = oldGroupInfo.DisplayName,
                    Description = oldGroupInfo.Description,
                    MailNickname = oldGroupInfo.MailNickname,
                    SiteUrl = oldGroupInfo.SiteUrl
                };
            }

            // update Office Group
            var accessToken = AuthHelper.GetMSGraphAccessToken();
            CreateOfficeGroup.FnLog.Info($"GroupId: {newGroupInfo.GroupId} - Personvern: {newGroupInfo.IsPrivate}");

            var isUpdated = OfficeGroupUtility.UpdateUnifiedGroup(
                newGroupInfo.GroupId,
                accessToken,
                10, 500,
                newGroupInfo.DisplayName != oldGroupInfo.DisplayName
                    ? newGroupInfo.DisplayName
                    : newGroupInfo.DisplayName,
                newGroupInfo.Description != oldGroupInfo.Description
                    ? newGroupInfo.Description
                    : newGroupInfo.Description,
                null, // members and owners will be dealt separately
                null, // members and owners will be dealt separately
                null,
                newGroupInfo.IsPrivate
            );

            // Update Modern Site if group info changed
            if (isUpdated && (newGroupInfo.DisplayName != oldGroupInfo.DisplayName ||
                              newGroupInfo.Description != oldGroupInfo.Description))
            {
                var gCtx = AuthHelper.GetAppOnlyAuthenticatedContext(oldGroupInfo.SiteUrl);
                var groupWeb = gCtx.Web;
                if (newGroupInfo.DisplayName != oldGroupInfo.DisplayName)
                    groupWeb.Title = newGroupInfo.DisplayName;

                if (newGroupInfo.Description != oldGroupInfo.Description)
                    groupWeb.Description = newGroupInfo.Description;

                groupWeb.Update();
                gCtx.ExecuteQuery();
            }

            handleGroupMembershipChanges(oldGroupInfo, newGroupInfo, oldGroup);

            // return updated Group info
            return new UnifiedGroupEntity
            {
                GroupId = newGroupInfo.GroupId,
                DisplayName = newGroupInfo.DisplayName,
                Description = newGroupInfo.Description,
                MailNickname = newGroupInfo.MailNickname,

                // SiteUrl does not change
                SiteUrl = oldGroupInfo.SiteUrl
            };
        }

        private static void handleGroupMembershipChanges(OfficeGroupInfo oldGroupInfo, OfficeGroupInfo newGroupInfo,
            Group oldGroup)
        {
            /**
                Adding users process
                    1. Get Owners and Members
                    2. Compare with /Members
                    3. Register the difference as /Members
                    4. Add Owners to /Members
             *
             * */
            // Work with members
            List<string> newMembers = new List<string>();
            List<string> removedMembers = new List<string>();
            List<string> newOwners = new List<string>();
            List<string> removedOwners = new List<string>();
            List<string> newMembersAndOwners = new List<string>();

            newMembersAndOwners = newGroupInfo.Members.Concat(newGroupInfo.Owners).ToList<string>();
            newMembers = newGroupInfo.Members.Except(oldGroupInfo.Members).ToList<string>();
            removedMembers = oldGroupInfo.Members.Except(newMembersAndOwners).ToList<string>();

            newOwners = newGroupInfo.Owners.Except(oldGroupInfo.Owners).ToList<string>();
            removedOwners = oldGroupInfo.Owners.Except(newGroupInfo.Owners).ToList<string>();

            // Process 3. Register the difference as /Members
            var allMembers = newGroupInfo.Members
                .Concat(newGroupInfo.Owners)
                .Concat(oldGroupInfo.Owners)
                .Distinct()
                .ToArray<string>();

            var allNewMembers = allMembers.Except(oldGroupInfo.Members).ToArray<string>();
            var allRemovedMembers = removedMembers.Concat(removedOwners).ToArray<string>();

            var graphClient = GraphHelper.GetAuthenticatedClient();

            /***
             *
                Removing:
                    Get Owners and Members,
                    compare with already registered users,
                    remove the difference from both /Owners and /Members
             */
            try
            {
                CreateOfficeGroup.FnLog.Info($"Group: {oldGroup.Id} - removing members");
                OfficeGroupUtility.RemoveMembers(allRemovedMembers.Distinct().ToArray<string>(), graphClient, oldGroup)
                    .GetAwaiter().GetResult();
            }
            catch (Exception e)
            {
                CreateOfficeGroup.FnLog.Error($"Error Group: {oldGroup.Id} removing members", e);
            }

            try
            {
                CreateOfficeGroup.FnLog.Info($"Group: {oldGroup.Id} - removing members");
                OfficeGroupUtility.RemoveOwners(removedOwners.ToArray<string>(), graphClient, oldGroup)
                    .GetAwaiter().GetResult();
            }
            catch (Exception e)
            {
                CreateOfficeGroup.FnLog.Error($"Error Group: {oldGroup.Id} removing owners", e);
            }

            // ADd all users as Members
            try
            {
                CreateOfficeGroup.FnLog.Info($"Group: {oldGroup.Id} - Adding all members");
                OfficeGroupUtility.AddMembers(allNewMembers.Distinct().ToArray<string>(), graphClient, oldGroup)
                    .GetAwaiter().GetResult();
            }
            catch (Exception e)
            {
                CreateOfficeGroup.FnLog.Error($"Error Group: {oldGroup.Id} adding all members", e);
            }

            // Add new Owners
            try
            {
                CreateOfficeGroup.FnLog.Info($"Group: {oldGroup.Id} - Adding owners");
                OfficeGroupUtility.AddOwners(newOwners.ToArray<string>(), graphClient, oldGroup)
                    .GetAwaiter().GetResult();
            }
            catch (Exception e)
            {
                CreateOfficeGroup.FnLog.Error($"Error Group: {oldGroup.Id} Adding owners", e);
            }
        }

        private static ListItem SetAuthorAsOtherRole(ListItem listItem, ClientContext clientContext, string role)
        {
            var listItemId = listItem.Id;

            FieldUserValue[] owners = new FieldUserValue[1];
            FieldUserValue authorsValue = (FieldUserValue) listItem["Author"];
            Microsoft.SharePoint.Client.User principal = clientContext.Web.GetUserById(authorsValue.LookupId);

            clientContext.Load(principal);
            clientContext.ExecuteQuery();

            owners[0] = new FieldUserValue();
            owners[0].LookupId = principal.Id;

            listItem[role] = owners;
            listItem.SystemUpdate();
            clientContext.Load(listItem);
            clientContext.ExecuteQuery();
            CreateOfficeGroup.FnLog.Info(
                $"OrderId: {listItemId} - Set Author as {role}, because {role} field was empty");

            return listItem;
        }

        private static OfficeGroupInfo ListItemToGroupInfo(ListItem listItem)
        {
            return new OfficeGroupInfo
            {
                DisplayName = listItem["Title"] as string,
                Description = listItem["NrkDescription"] as string,
                GroupType = listItem["NrkGroupType"] as string == "Adhoc" ? GroupType.AdHoc : GroupType.Structured,
                IsPrivate = (listItem["NrkPrivacy"] as string).ToLower() == "Privat".ToLower() ? true : false,
                MailNickname = listItem["Title"] as string,
                Members = GetUserFieldEmails(listItem, "NrkMembers"),
                Owners = GetUserFieldEmails(listItem, "NrkGroupOwner"),
                GroupId = listItem["NrkGroupId"] as string,
                SiteUrl = listItem["NrkGroupUrl"] as string
            };
        }

        private static async Task<OfficeGroupInfo> GetGroupInfoFromUnifiedGroup(Group unifiedGroup, string accessToken)
        {
            var members =
                await OfficeGroupManager.GetGroupUsers(unifiedGroup.Id, OfficeGroupManager.GroupUserEndpoint.Members,
                    accessToken);
            var owners =
                await OfficeGroupManager.GetGroupUsers(unifiedGroup.Id, OfficeGroupManager.GroupUserEndpoint.Owners,
                    accessToken);

            //var members = unifiedGroup.Members[0]
            return new OfficeGroupInfo
            {
                DisplayName = unifiedGroup.DisplayName,
                Description = unifiedGroup.Description,
                GroupType = GroupType.AdHoc | GroupType.Structured,
                IsPrivate = unifiedGroup.Visibility == "Public" ? false : true,
                MailNickname = unifiedGroup.MailNickname,
                Members = members.ToArray(),
                Owners = owners.ToArray(),
                GroupId = unifiedGroup.Id,
                SiteUrl = OfficeGroupUtility.GetUnifiedGroupSiteUrl(unifiedGroup.Id, AuthHelper.GetMSGraphAccessToken())
            };
        }

        private static string[] GetUserFieldEmails(ListItem listItem, string spFieldName)
        {
            List<string> members = new List<string>();
            foreach (FieldUserValue userValue in listItem[spFieldName] as FieldUserValue[])
            {
                members.Add(userValue.Email);
            }

            return members.ToArray();
        }
    }
}
