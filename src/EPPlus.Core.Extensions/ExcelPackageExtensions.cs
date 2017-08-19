using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Core.Extensions
{
    /// <summary>
    /// Class holds extensions on ExcelPackage object
    /// </summary>
    public static class ExcelPackageExtensions
    {
        /// <summary>
        /// Method returns all table names in the opened worksheet
        /// </summary>
        /// <remarks>Excel is ensuring the uniqueness of table names</remarks>
        /// <param name="excel">Extended ExcelPackage object</param>
        /// <returns>Enumeration of ExcelTables</returns>
        public static IEnumerable<ExcelTable> GetTables(this ExcelPackage excel)
        {
            foreach (ExcelWorksheet ws in excel.Workbook.Worksheets)
            {
                foreach (ExcelTable t in ws.Tables)
                    yield return t;
            }
        }

        /// <summary>
        /// Method returns concrete ExcelTable by it's name 
        /// </summary>
        /// <param name="excel">Extended ExcelPackage object</param>
        /// <param name="name">Name of the table</param>
        /// <returns>ExcelTable object if found, null inf not</returns>
        public static ExcelTable GetTable(this ExcelPackage excel, string name)
        {
            return excel.GetTables().FirstOrDefault(t => t.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Method checks for table in the ExcelPackage
        /// </summary>
        /// <param name="excel">Extended ExcelPackage object</param>
        /// <param name="name">Name of the table</param>
        /// <returns>Result of search as bool</returns>
        public static bool HasTable(this ExcelPackage excel, string name)
        {
            return excel.GetTables().Any(t => t.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
        }
    }
}
