using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using Office365Groups.Library.ModelValidation;

namespace Office365Groups.Library
{
	public class MetadataInfo
	{																		 
		[Required(ErrorMessage = "DisplayName is required")]
		public string DisplayName { get; set; }

        [Required(ErrorMessage = "GroupId is required")]
        public string GroupId { get; set; }

		public string Description { get; set; }

        [Required(ErrorMessage = "SiteType is required")]
        public string SiteType { get; set; }

		[Required]
		[EmailArrayValidation(ErrorMessage = "At least one owner is required and it must be a valid email address.")]
		public string[] Owners { get; set; }

        public string ParentDepartment { get; set; }

        public string OwnedDepartment { get; set; }

        public string ProjectNumber { get; set; }

        public string ShortName { get; set; }

        public string ProjectGoal { get; set; }

        public string ProjectPurpose { get; set; }

        public DateTime StartDate { get; set; }
        
        public DateTime EndDate { get; set; }

        public string OwnerDisplayName { get; set; }

        public string ShortDisplayName { get; set; }

		public void Validate()
		{
			var result = CustomValidations.Validate(this);
			if (result.isValid == false)
			{
				var error = result.ValidationExceptions.FirstOrDefault();
				throw new ValidationException(error.Message, error);
			}
		}

	}
}
