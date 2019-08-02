using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace Office365Groups.Library.ModelValidation
{
	[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field | AttributeTargets.Parameter, AllowMultiple = false)]
	public class EmailArrayValidation: ValidationAttribute
	{
		protected override ValidationResult IsValid(object value, ValidationContext validationContext)
		{
			string[] array = value as string[];

			if (array != null)
			{
				//// if empty not valid
				if (array.Length == 0)
					return new ValidationResult(string.Format("{0} field requires at least one valid email address.", validationContext.DisplayName));

				Console.WriteLine(ErrorMessage);
				EmailAddressAttribute emailAttribute = new EmailAddressAttribute();
				
				foreach (string str in array)
				{
					//// if all are not valid emails, then not valid
					if (!emailAttribute.IsValid(str))
					{
						return new ValidationResult(string.Format("{0} field values must be valid email addresses. {1} is not valid email address.", validationContext.DisplayName, str));
					}
				}

				return ValidationResult.Success;
			}

			return base.IsValid(value, validationContext);
		}

		public override string FormatErrorMessage(string name)
		{
			return String.Format(ErrorMessage, name);
		}


	}
}
