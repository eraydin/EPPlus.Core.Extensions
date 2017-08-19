using System;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions
{
    /// <summary>
    /// Class contains exception circumstances
    /// </summary>
    public class ExcelTableConvertExceptionArgs
    {
        /// <summary>
        /// Property that was tried to set
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// Column that was mapped to this property
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// Type of the property
        /// </summary>
        public Type ExpectedType { get; set; }

        /// <summary>
        /// Cell value returned by EPPlus
        /// </summary>
        public object CellValue { get; set; }

        /// <summary>
        /// Absolute address of the cell, where the conversion error occured
        /// </summary>
        public ExcelCellAddress CellAddress { get; set; }
    }

    /// <summary>
    /// Class extends exception to hold casting exception circumstances
    /// </summary>
    public class ExcelTableConvertException : Exception
    {
        public ExcelTableConvertExceptionArgs Args { get; private set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public ExcelTableConvertException()
        {
        }

        /// <summary>
        /// Constructor with message
        /// </summary>
        /// <param name="message">Exception message</param>
        public ExcelTableConvertException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Constructor with message and inner exception
        /// </summary>
        /// <param name="message">Exception message</param>
        /// <param name="inner">Inner exception</param>
        public ExcelTableConvertException(string message, Exception inner)
            : base(message, inner)
        {
        }

        /// <summary>
        /// Custom constructor that takes message,, inner exception and conversion arguments
        /// </summary>
        /// <param name="message">Exception message</param>
        /// <param name="inner">Actual conversion exception catched</param>
        /// <param name="args">Information related to the circumstances of the exception</param>
        public ExcelTableConvertException(string message, Exception inner, ExcelTableConvertExceptionArgs args)
            : base(message, inner)
        {
            this.Args = args;
        }
    }
}
