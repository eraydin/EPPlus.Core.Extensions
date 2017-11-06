using System;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions.Configuration
{
    /// <summary>
    ///     Default configurations
    /// </summary>
    internal class DefaultExcelConfiguration<T> : IExcelConfiguration<T>
    {
        public static IExcelConfiguration<T> Instance => new DefaultExcelConfiguration<T>();

        public bool HasHeaderRow { get; set; } = true;

        public bool SkipCastingErrors { get; set; } = false;

        public bool SkipValidationErrors { get; set; } = false;

        public Action<ExcelRange, T> ConfigureCell { get; set; } = null;

        public Action<ExcelColumn> ConfigureColumn { get; set; } = null;

        public Action<ExcelRange> ConfigureHeader { get; set; } = null;

        public Action<ExcelRange> ConfigureHeaderRow { get; set; } = null;
    }
}
