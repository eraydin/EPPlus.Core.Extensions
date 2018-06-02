using System;

namespace EPPlus.Core.Extensions
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelWorksheetAttribute : Attribute
    {
        public ExcelWorksheetAttribute()
        {
        }

        public ExcelWorksheetAttribute(string worksheetName) => WorksheetName = worksheetName;

        public string WorksheetName { get; protected set; }
    }
}
