using System;

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace EPPlus.Core.Extensions.Configuration
{
    /// <summary>
    ///     Default configurations
    /// </summary>
    public class DefaultExcelConfiguration : IExcelConfiguration
    {
        public bool HasHeaderRow { get; set; } = true;

        public bool SkipCastingErrors { get; set; } = false;

        public bool SkipValidationErrors { get; set; } = false;

        public Action<ExcelRange, T> CellConfiguration { get; set; } = null;

        public Action<ExcelColumn> ColumnConfiguration { get; set; } = null;

        public Action<ExcelRange> HeaderConfiguration { get; set; } = null;

        public Action<ExcelRange> HeaderRowConfiguration { get; set; } = null;
    }
}
