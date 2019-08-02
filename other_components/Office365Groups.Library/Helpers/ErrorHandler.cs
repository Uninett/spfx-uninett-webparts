using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Office365Groups.Library
{
	public static class ErrorHandler
	{
		public static HttpResponseMessage CreateExceptionResponse(HttpRequestMessage req, Exception ex)
		{
			if (ex is ArgumentNullException || ex is ArgumentException || ex is ValidationException)
				return req.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message, ex);

			// Microsoft Graph error message is in .Error.Message
			if (ex is Microsoft.Graph.ServiceException)
			{
				var graphError = (Microsoft.Graph.ServiceException)ex;
				return req.CreateErrorResponse(HttpStatusCode.InternalServerError, graphError.Error.Message, ex);
			}

            // Handle Inner exceptions
            if (ex.InnerException != null)
            {
                // Deserialize exception if it is Microsoft.SharePoint.SPException
                if (ex.InnerException.Message.IndexOf("Microsoft.SharePoint.SPException") >= 0)
                {
                    var result = JsonConvert.DeserializeObject<SPExceptionMessage>(ex.InnerException.Message);
                    
                    return req.CreateErrorResponse(HttpStatusCode.InternalServerError, result.error.message, ex);
                }
 
                return req.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.InnerException.Message, ex);
            }


            return req.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message, ex);

		}

	}
}
