using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Office365Groups.Library.Helpers
{
	public static class ManageSPFxExtensions
	{
		// TODO: provide details of a single site here, or provide some code to iterate several sites..
		//private static readonly string SITE_URL = "TODO";
		//private static readonly string USERNAME = "TODO";
		//private static readonly string PWD = "TODO";
		private static Guid TEMPLATAEHUB_COMPONENT_ID = new Guid("997fe955-efac-4f36-8383-e51c3c8eed40");

		public static void AddTemplateHub(string siteUrl)
		{
			string userName = AuthHelper.SPUserId;
			string password = AuthHelper.SPUserPassword;

			AuthenticationManager authMgr = new AuthenticationManager();
			ClientContext ctx = authMgr.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, userName, password);
			Web currentWeb = ctx.Web;
			ctx.Load(currentWeb);
			ctx.ExecuteQueryRetry();
			
			// TODO: comment in/out method calls here as you need..
			var customActions = getCustomActions(ctx, TEMPLATAEHUB_COMPONENT_ID);
			//addSpfxExtensionCustomAction(ctx);
			//removeSpfxExtensionCustomAction(ctx, "COB-SPFx-GlobalHeader");
			
			if (customActions != null && customActions.Count > 0)
			{
				return;
			}

			string spfxExtName = "New from TemplateHub Spfx";
			string spfxExtTitle = "New from TemplateHub Spfx";
			string spfxExtGroup = "";
			string spfxExtDescription = "Create new document from TemplateHub";
			string spfxExtLocation = "ClientSideExtension.ListViewCommandSet";

			BasePermissions basePermission = new BasePermissions() ;
			basePermission.Set(PermissionKind.EditListItems);
			CustomActionEntity ca = new CustomActionEntity
			{
				Name = spfxExtName,
				RegistrationId = "101",
				RegistrationType = UserCustomActionRegistrationType.List,
				Title = spfxExtTitle,
				Group = spfxExtGroup,
				Description = spfxExtDescription,
				Rights = basePermission,
				Sequence = 1,
				Location = spfxExtLocation,
				ClientSideComponentId = TEMPLATAEHUB_COMPONENT_ID
			};

			addSpfxExtensionCustomAction(ctx, ca);
		}


		private static void addSpfxExtensionCustomAction(ClientContext ctx, CustomActionEntity spfxCustomAction)
		{
			ctx.Web.AddCustomAction(spfxCustomAction);
			ctx.ExecuteQueryRetry();
		}

		private static void removeSpfxExtensionCustomAction(ClientContext ctx, string extensionName, string location = "ClientSideExtension.ApplicationCustomizer")
		{
			Console.WriteLine(string.Format("removeSpfxExtensionCustomAction: Removing extension with name '{0}' from web - '{1}'.",
					extensionName, ctx.Web.Url));

			var customActionsToRemove = ctx.Web.GetCustomActions().Where(ca =>
					ca.Location == location &&
					ca.Name == extensionName);

			if (customActionsToRemove.Count() > 0)
			{
				foreach (var ca in customActionsToRemove)
				{
					ctx.Web.DeleteCustomAction(ca.Id);
					ctx.Web.Update();
					ctx.ExecuteQueryRetry();

					Console.WriteLine(string.Format("removeSpfxExtensionCustomAction: Successfully removed extension '{0}' from web '{1}'.",
							extensionName, ctx.Web.Url));
				}
			}
			else
			{
				Console.WriteLine(string.Format("removeSpfxExtensionCustomAction: Did not find any extension to remove named '{0}' on web '{1}'.",
								extensionName, ctx.Web.Url));
			}

			Console.WriteLine("removeSpfxExtensionCustomAction: Leaving.");
		}

		public static List<UserCustomAction> getCustomActions(ClientContext ctx, Guid componentId)
		{
			Console.WriteLine(string.Format("getCustomActions: Retrieving extensions for web - '{0}'.", ctx.Web.Url));

			var customActions = ctx.Web.GetCustomActions().Where(ca => ca.ClientSideComponentId == componentId);
			if (customActions.Count() > 0)
			{
				return customActions.ToList() ;
			}
			else
			{
				return null;
			}
		}

	}
}
