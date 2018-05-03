using System;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions.Configuration
{
    public interface IExcelCreateConfiguration<T>
    {
        Action<ExcelRange> ConfigureTitle { get; }

        Action<ExcelRange, T> ConfigureCell { get; }

        Action<ExcelColumn> ConfigureColumn { get; }

        Action<ExcelRange> ConfigureHeader { get; }

        Action<ExcelRange> ConfigureHeaderRow { get; }

        IExcelCreateConfiguration<T> WithCellConfiguration(Action<ExcelRange, T> cellConfiguration);

        IExcelCreateConfiguration<T> WithColumnConfiguration(Action<ExcelColumn> columnConfiguration);

        IExcelCreateConfiguration<T> WithHeaderConfiguration(Action<ExcelRange> headerConfiguration);

        IExcelCreateConfiguration<T> WithHeaderRowConfiguration(Action<ExcelRange> headerRowConfiguration);

        IExcelCreateConfiguration<T> WithTitleConfiguration(Action<ExcelRange> titleConfiguration);
    }

    public class DefaultExcelCreateConfiguration<T> : IExcelCreateConfiguration<T>
    {
        public static IExcelCreateConfiguration<T> Instance => new DefaultExcelCreateConfiguration<T>();

        public Action<ExcelRange> ConfigureTitle { get; protected set; }

        public Action<ExcelRange, T> ConfigureCell { get; protected set; }

        public Action<ExcelColumn> ConfigureColumn { get; protected set; }

        public Action<ExcelRange> ConfigureHeader { get; protected set; }

        public Action<ExcelRange> ConfigureHeaderRow { get; protected set; }

        public IExcelCreateConfiguration<T> WithCellConfiguration(Action<ExcelRange, T> cellConfiguration)
        {
            ConfigureCell = cellConfiguration;
            return this;
        }

        public IExcelCreateConfiguration<T> WithColumnConfiguration(Action<ExcelColumn> columnConfiguration)
        {
            ConfigureColumn = columnConfiguration;
            return this;
        }

        public IExcelCreateConfiguration<T> WithHeaderConfiguration(Action<ExcelRange> headerConfiguration)
        {
            ConfigureHeader = headerConfiguration;
            return this;
        }

        public IExcelCreateConfiguration<T> WithHeaderRowConfiguration(Action<ExcelRange> headerRowConfiguration)
        {
            ConfigureHeaderRow = headerRowConfiguration;
            return this;
        }

        public IExcelCreateConfiguration<T> WithTitleConfiguration(Action<ExcelRange> titleConfiguration)
        {
            ConfigureTitle = titleConfiguration;
            return this;
        }
    }
}
