using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office365Groups.Library.Helpers
{
    public class TenantSharingCapibilityException: Exception
    {
        public TenantSharingCapibilityException()
        {
        }

        public TenantSharingCapibilityException(string message)
            : base(message)
        {
        }

        public TenantSharingCapibilityException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
