using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EPPlus.Core.Extensions.Validation
{
    [Serializable]
    internal class EntityValidationResult
    {
        public EntityValidationResult(IList<ValidationResult> violations = null)
        {
            ValidationErrors = violations ?? new List<ValidationResult>();
        }

        public IList<ValidationResult> ValidationErrors { get; }

        public bool HasError => ValidationErrors.Count > 0;
    }
}
