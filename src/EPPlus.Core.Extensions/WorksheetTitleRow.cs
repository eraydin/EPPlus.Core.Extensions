using OfficeOpenXml;
using System;

namespace EPPlus.Core.Extensions
{
    internal class WorksheetTitleRow
    {
        internal string Title { get; set; }
        internal Action<ExcelRange> ConfigureTitle { get; set; }
    }
}