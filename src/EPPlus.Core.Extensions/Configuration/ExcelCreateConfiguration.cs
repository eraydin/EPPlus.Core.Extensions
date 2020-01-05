using System;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions.Configuration
{
    public class ExcelCreateConfiguration<T>
    {
        public static ExcelCreateConfiguration<T> Instance => new ExcelCreateConfiguration<T>();

        public virtual Action<ExcelRange> ConfigureTitle { get; private set; }

        public virtual Action<ExcelRange, T> ConfigureCell { get; private set; }

        public virtual Action<ExcelColumn> ConfigureColumn { get; private set; }

        public virtual Action<ExcelRange> ConfigureHeader { get; private set; }

        public virtual Action<ExcelRange> ConfigureHeaderRow { get; private set; }

        public virtual ExcelCreateConfiguration<T> WithCellConfiguration(Action<ExcelRange, T> cellConfiguration)
        {
            ConfigureCell = cellConfiguration;
            return this;
        }

        public virtual ExcelCreateConfiguration<T> WithColumnConfiguration(Action<ExcelColumn> columnConfiguration)
        {
            ConfigureColumn = columnConfiguration;
            return this;
        }

        public virtual ExcelCreateConfiguration<T> WithHeaderConfiguration(Action<ExcelRange> headerConfiguration)
        {
            ConfigureHeader = headerConfiguration;
            return this;
        }

        public virtual ExcelCreateConfiguration<T> WithHeaderRowConfiguration(Action<ExcelRange> headerRowConfiguration)
        {
            ConfigureHeaderRow = headerRowConfiguration;
            return this;
        }

        public virtual ExcelCreateConfiguration<T> WithTitleConfiguration(Action<ExcelRange> titleConfiguration)
        {
            ConfigureTitle = titleConfiguration;
            return this;
        }
    }
}        