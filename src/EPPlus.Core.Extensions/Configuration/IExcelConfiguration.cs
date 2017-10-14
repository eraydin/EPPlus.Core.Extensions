using System;

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace EPPlus.Core.Extensions.Configuration
{
    public interface IExcelConfiguration
    {
        bool HasHeaderRow { get; set; }

        /// <summary>
        ///     Determines how the method should handle exceptions when casting cell value to property type.
        ///     If this is true, invalid casts are silently skipped, otherwise any error will cause method to fail with exception.
        /// </summary>
        bool SkipCastingErrors { get; set; }

        bool SkipValidationErrors { get; set; }

        Action<ExcelRange, T> CellConfiguration { get; set; }

        Action<ExcelColumn> ColumnConfiguration { get; set; }

        Action<ExcelRange> HeaderConfiguration { get; set; }

        Action<ExcelRange> HeaderRowConfiguration { get; set; }
    }
}
