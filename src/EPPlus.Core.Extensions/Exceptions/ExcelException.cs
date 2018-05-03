using System;

namespace EPPlus.Core.Extensions.Exceptions
{
    /// <inheritdoc />
    /// <summary>
    ///     Class extends exception to hold casting exception circumstances
    /// </summary>
    public class ExcelException : Exception
    {
        /// <inheritdoc />
        /// <summary>
        ///     Default constructor
        /// </summary>
        protected ExcelException()
        {
        }

        /// <inheritdoc />
        /// <summary>
        ///     Constructor with message
        /// </summary>
        /// <param name="message">Exception message</param>
        public ExcelException(string message)
            : base(message)
        {
        }

        /// <inheritdoc />
        /// <summary>
        ///     Constructor with message and inner exception
        /// </summary>
        /// <param name="message">Exception message</param>
        /// <param name="inner">Inner exception</param>
        public ExcelException(string message, Exception inner)
            : base(message, inner)
        {
        }

        public ExcelExceptionArgs Args { get; protected set; }

        public ExcelException WithArguments(ExcelExceptionArgs args)
        {
            Args = args;
            return this;
        }
    }
}
