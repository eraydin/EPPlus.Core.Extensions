using System;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions.Configuration
{
    public interface IExcelConfiguration<T>
    {
        bool HasHeaderRow { get; set; }

        /// <summary>
        ///     Determines how the method should handle exceptions when casting cell value to property type.
        ///     If this is true, invalid casts are silently skipped, otherwise any error will cause method to fail with exception.
        /// </summary>
        bool SkipCastingErrors { get; set; }

        bool SkipValidationErrors { get; set; }

        Action<ExcelRange, T> ConfigureCell { get; set; }

        Action<ExcelColumn> ConfigureColumn { get; set; }

        Action<ExcelRange> ConfigureHeader { get; set; }

        Action<ExcelRange> ConfigureHeaderRow { get; set; }
    }
}
