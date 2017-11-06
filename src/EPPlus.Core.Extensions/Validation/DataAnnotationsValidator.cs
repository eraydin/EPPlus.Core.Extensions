using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EPPlus.Core.Extensions.Validation
{
    internal static class DataAnnotationsValidator
    {
        public static void Validate<T>(this T thisObject, int rowIndex)
        {
            var validationResults = new List<ValidationResult>();
            var validationContext = new ValidationContext(thisObject, null, null);
            bool isValid = Validator.TryValidateObject(thisObject, validationContext, validationResults, true);

            if (!isValid)
            {
                throw new ExcelTableValidationException($"Validation failed on the {rowIndex}. row of ExcelTable. See 'ValidationErrors' property for more details.", validationResults);
            }
        }
    }
}
