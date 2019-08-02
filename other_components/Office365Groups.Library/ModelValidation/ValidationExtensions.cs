using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.ComponentModel.DataAnnotations;

namespace Office365Groups.Library.ModelValidation
{
	public static class CustomValidations
	{
		public static bool ValidateModel<T>(this string json) where T : new()
		{
			T model = new JavaScriptSerializer().Deserialize<T>(json);
			return ValidateModel<T>(model);
		}

		public static bool ValidateModel<T>(this T model) where T : new()
		{
			var validationContext = new ValidationContext(model, null, null);
			return Validator.TryValidateObject(model, validationContext, null, true);
		}

		public static (bool isValid, List<ValidationException> ValidationExceptions) Validate(object model)
		{
			ValidationContext validationContext = new ValidationContext(model, null, null);
			List<ValidationResult> validationResults = new List<ValidationResult>();
			bool isValid = Validator.TryValidateObject(model, validationContext, validationResults, true);

			List<ValidationException> validationExceptions = new List<ValidationException>();
			if (!isValid)
			{
				foreach (var validationResult in validationResults)
				{
					var error = new ValidationException(validationResult, null, model);
					validationExceptions.Add(error);
				}
			}

			return (isValid, validationExceptions);
		}

	}

}
