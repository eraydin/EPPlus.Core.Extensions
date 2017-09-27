using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EPPlus.Core.Extensions.Validation
{
    internal class EntityValidator<T> where T : class
    {
        public EntityValidationResult Validate(T entity)
        {
            var validationResults = new List<ValidationResult>();
            var validationContext = new ValidationContext(entity, null, null);
            bool isValid = Validator.TryValidateObject(entity, validationContext, validationResults, true);

            return new EntityValidationResult(validationResults);
        }
    }
}
