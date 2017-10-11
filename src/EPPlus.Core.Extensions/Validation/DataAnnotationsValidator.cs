using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EPPlus.Core.Extensions.Validation
{
    internal static class DataAnnotationsValidator
    {
        public static void Validate<T>(this T thisObject)
        {
            var validationResults = new List<ValidationResult>();
            var validationContext = new ValidationContext(thisObject, null, null);
            bool isValid = Validator.TryValidateObject(thisObject, validationContext, validationResults, true);

            if (!isValid)
            {
                throw new ExcelTableValidationException("Validation failed for one or more objects. See 'ValidationErrors' property for more details.", validationResults);
            }
        }
    }
}
