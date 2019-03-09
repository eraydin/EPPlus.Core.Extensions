using System;

namespace EPPlus.Core.Extensions.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelWorksheetAttribute : Attribute
    {
        public ExcelWorksheetAttribute()
        {
        }

        public ExcelWorksheetAttribute(string worksheetName) => WorksheetName = worksheetName;

        public string WorksheetName { get; }
    }
}
