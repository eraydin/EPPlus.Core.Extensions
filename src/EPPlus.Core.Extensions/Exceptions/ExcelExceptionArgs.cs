using System;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions.Exceptions
{
    /// <summary>
    ///     Class contains exception circumstances
    /// </summary>
    public class ExcelExceptionArgs
    {
        /// <summary>
        ///     Property that was tried to set
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        ///     Column that was mapped to this property
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        ///     Type of the property
        /// </summary>
        public Type ExpectedType { get; set; }

        /// <summary>
        ///     Cell value returned by EPPlus
        /// </summary>
        public object CellValue { get; set; }

        /// <summary>
        ///     Absolute address of the cell, where the conversion error occured
        /// </summary>
        public ExcelCellAddress CellAddress { get; set; }
    }
}
