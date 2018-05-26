using System;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions.Configuration
{
    public abstract class ExcelCreateConfiguration<T>
    {
        internal Action<ExcelRange> ConfigureTitle { get; private set; }

        internal Action<ExcelRange, T> ConfigureCell { get; private set; }

        internal Action<ExcelColumn> ConfigureColumn { get; private set; }

        internal Action<ExcelRange> ConfigureHeader { get; private set; }

        internal Action<ExcelRange> ConfigureHeaderRow { get; private set; }

        public ExcelCreateConfiguration<T> WithCellConfiguration(Action<ExcelRange, T> cellConfiguration)
        {
            ConfigureCell = cellConfiguration;
            return this;
        }

        public ExcelCreateConfiguration<T> WithColumnConfiguration(Action<ExcelColumn> columnConfiguration)
        {
            ConfigureColumn = columnConfiguration;
            return this;
        }

        public ExcelCreateConfiguration<T> WithHeaderConfiguration(Action<ExcelRange> headerConfiguration)
        {
            ConfigureHeader = headerConfiguration;
            return this;
        }

        public ExcelCreateConfiguration<T> WithHeaderRowConfiguration(Action<ExcelRange> headerRowConfiguration)
        {
            ConfigureHeaderRow = headerRowConfiguration;
            return this;
        }

        public ExcelCreateConfiguration<T> WithTitleConfiguration(Action<ExcelRange> titleConfiguration)
        {
            ConfigureTitle = titleConfiguration;
            return this;
        }
    }

    public class DefaultExcelCreateConfiguration<T> : ExcelCreateConfiguration<T>
    {
        public static ExcelCreateConfiguration<T> Instance => new DefaultExcelCreateConfiguration<T>();
    }
}        