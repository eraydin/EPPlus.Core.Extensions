using System;

using EPPlus.Core.Extensions.Exceptions;

namespace EPPlus.Core.Extensions.Results
{
    /// <summary>
    ///     Describes an error captured while importing an Excel row.
    /// </summary>
    public sealed class ExcelReadError
    {
        internal ExcelReadError(ExcelReadErrorKind kind, string message, Exception exception, ExcelExceptionArgs context)
        {
            Kind = kind;
            Message = message;
            Exception = exception;
            Context = context;
        }

        public ExcelReadErrorKind Kind { get; }

        public string Message { get; }

        public Exception Exception { get; }

        public ExcelExceptionArgs Context { get; }
    }
}
