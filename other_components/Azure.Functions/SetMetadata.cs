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
using Office365Groups.Library.Models;

namespace Rederi.Functions
{
    public static class SetMetadata
    {
        public static TraceWriter FnLog { get; set; }

        [FunctionName("SetMetadata")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]HttpRequestMessage req, ExecutionContext context, TraceWriter log)
        {
            FnLog = log;
            OfficeGroupUtility.Logger = FnLog;

            var metaReq = await OfficeGroupManager.ExtractMetadata(req);
            metaReq.Validate();

            var groupId = metaReq.GroupId;
            var metadata = new InmetaGenericExtension(metaReq);

            try
            {
                await OfficeGroupUtility.UpdateGroupMetaData(groupId, metadata, "extvcs569it_InmetaGenericSchema");
                return req.CreateResponse(HttpStatusCode.OK, groupId);
            } catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                return ErrorHandler.CreateExceptionResponse(req, ex);
            }

        }

    }
}