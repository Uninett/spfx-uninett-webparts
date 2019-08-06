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
	public class SPExceptionMessage
	{
        public class SPExceptionMessageError
        {
            public string code { get; set; }
            public string message { get; set; }
        }
        public SPExceptionMessageError error { get; set; }
    }
}
