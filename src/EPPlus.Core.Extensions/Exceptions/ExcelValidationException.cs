using System;

namespace EPPlus.Core.Extensions.Exceptions
{
    public class ExcelValidationException : ExcelException
    {
        public ExcelValidationException(string message)
            : base(message)
        {
        }

        public ExcelValidationException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
