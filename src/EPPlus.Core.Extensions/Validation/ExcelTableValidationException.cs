using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EPPlus.Core.Extensions.Validation
{
    public class ExcelTableValidationException : Exception
    {
        protected ExcelTableValidationException() { }

        public ExcelTableValidationException(string message, IEnumerable<ValidationResult> validationResults) : base(message)
        {
            ValidationErrors = validationResults;
        }

        public IEnumerable<ValidationResult> ValidationErrors { get; }

        public ExcelTableExceptionArgs Args { get; protected set;}

        internal ExcelTableValidationException AddExceptionArguments(ExcelTableExceptionArgs args)
        {
            Args = args;
            return this;
        }
    }
}
