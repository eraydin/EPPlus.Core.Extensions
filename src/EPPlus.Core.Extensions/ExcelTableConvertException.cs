using System;

namespace EPPlus.Core.Extensions
{
    /// <inheritdoc />
    /// <summary>
    ///     Class extends exception to hold casting exception circumstances
    /// </summary>
    public class ExcelTableConvertException : Exception
    {
        /// <inheritdoc />
        /// <summary>
        ///     Default constructor
        /// </summary>
        protected ExcelTableConvertException() { }

        /// <inheritdoc />
        /// <summary>
        ///     Constructor with message
        /// </summary>
        /// <param name="message">Exception message</param>
        public ExcelTableConvertException(string message)
            : base(message) { }

        /// <inheritdoc />
        /// <summary>
        ///     Constructor with message and inner exception
        /// </summary>
        /// <param name="message">Exception message</param>
        /// <param name="inner">Inner exception</param>
        public ExcelTableConvertException(string message, Exception inner)
            : base(message, inner) { }

        /// <inheritdoc />
        /// <summary>
        ///     Custom constructor that takes message, inner exception and conversion arguments
        /// </summary>
        /// <param name="message">Exception message</param>
        /// <param name="inner">Actual conversion exception catched</param>
        /// <param name="args">Information related to the circumstances of the exception</param>
        public ExcelTableConvertException(string message, Exception inner, ExcelTableExceptionArgs args)
            : base(message, inner)
        {
            Args = args;
        }

        public ExcelTableExceptionArgs Args { get; }
    }
}
