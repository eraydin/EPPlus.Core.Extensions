using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EPPlus.Core.Extensions.Validation
{
    public class ExcelTableValidationException : Exception
    {
        protected ExcelTableValidationException() { }

        public ExcelTableValidationException(string message) : base(message)
        {
        }

        public ExcelTableValidationException(string message, IEnumerable<ValidationResult> validationResults) : base(message)
        {
            ValidationErrors = validationResults;
        }

        public IEnumerable<ValidationResult> ValidationErrors { get; }
    }
}
