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
	public class OfficeGroupInfo
	{
		[Required(ErrorMessage = "DisplayName is required")]
		public string DisplayName { get; set; }

		public string Description { get; set; }

		[Required]
		[EmailArrayValidation(ErrorMessage = "At least one owner is required and it must be a valid email address.")]
		public string[] Owners { get; set; }

		public string[] Members { get; set; }

		[JsonConverter(typeof(StringEnumConverter))]
		[Required(ErrorMessage = "GroupType is required.")]
		public GroupType GroupType { get; set; } = GroupType.AdHoc;

	    [Required(ErrorMessage = "MailNickname is required")]
        public string MailNickname { get; set; }

		[Required(ErrorMessage = "isPrivate is required")]
		public bool IsPrivate { get; set; } = false;

    public bool ExternalSharing { get; set; } = false;

    public bool CreateTeam { get; set; } = false;

		public string SiteUrl { get; set; }

		public string GroupId { get; set; }

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

	public enum GroupType
	{
		AdHoc,
		Structured
	}
}
