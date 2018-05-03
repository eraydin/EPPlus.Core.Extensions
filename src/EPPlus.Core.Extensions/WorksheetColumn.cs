using System;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions
{
    internal class WorksheetColumn<T>
    {
        internal Func<T, object> Map { get; set; }

        internal string Header { get; set; }

        internal Action<ExcelColumn> ConfigureColumn { get; set; }
    }
}
