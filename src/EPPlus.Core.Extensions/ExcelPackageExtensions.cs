using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace EPPlus.Core.Extensions
{
    /// <summary>
    ///     Class holds extensions on ExcelPackage object
    /// </summary>
    public static class ExcelPackageExtensions
    {
        /// <summary>
        ///     Returns all table names in the opened worksheet
        /// </summary>
        /// <remarks>Excel is ensuring the uniqueness of table names</remarks>
        /// <param name="excelPackage">The ExcelPackage object</param>
        /// <returns>Enumeration of ExcelTables</returns>
        public static IEnumerable<ExcelTable> GetTables(this ExcelPackage excelPackage)
        {
            foreach (ExcelWorksheet ws in excelPackage.Workbook.Worksheets)
            {
                foreach (ExcelTable t in ws.Tables)
                    yield return t;
            }
        }

        /// <summary>
        ///     Returns concrete ExcelTable by its name 
        /// </summary>
        /// <param name="excelPackage">The ExcelPackage object</param>
        /// <param name="name">Name of the table</param>
        /// <returns>ExcelTable object if found, null if not</returns>
        public static ExcelTable GetTable(this ExcelPackage excelPackage, string name)
        {
            return excelPackage.GetTables().FirstOrDefault(t => t.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        ///     Checks that given table name is in the ExcelPackage or not
        /// </summary>
        /// <param name="excel">The ExcelPackage object</param>
        /// <param name="name">Name of the table</param>
        /// <returns>Result of search as bool</returns>
        public static bool HasTable(this ExcelPackage excel, string name)
        {
            return excel.GetTables().Any(t => t.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        ///     Extracts a DataSet from the ExcelPackage.
        /// </summary>
        /// <param name="excelPackage">The ExcelPackage.</param>
        /// <param name="hasHeaderRow">Indicates whether worksheet has a header row or not.</param>
        /// <returns></returns>
        public static DataSet ToDataSet(this ExcelPackage excelPackage, bool hasHeaderRow = true)
        {
            var dataSet = new DataSet();

            foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
            {
                dataSet.Tables.Add(worksheet.ToDataTable(hasHeaderRow));
            }

            return dataSet;
        }

        /// <summary> 
        ///     Yields objects of specified type from given ExcelPackage
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excelPackage"></param>
        /// <param name="worksheetIndex"></param>
        /// <param name="skipCastingErrors"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns></returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelPackage excelPackage, int worksheetIndex = 1, bool skipCastingErrors = false, bool hasHeaderRow = true) where T : class, new()
        {
            return excelPackage.Workbook.Worksheets[worksheetIndex].AsEnumerable<T>(skipCastingErrors, hasHeaderRow);
        }

        /// <summary>
        ///     Converts given ExcelPackage to list of objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excelPackage"></param>
        /// <param name="worksheetIndex"></param>
        /// <param name="skipCastingErrors"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns></returns>
        public static IList<T> ToList<T>(this ExcelPackage excelPackage, int worksheetIndex = 1, bool skipCastingErrors = false, bool hasHeaderRow = true) where T : class, new()
        {
            return excelPackage.AsEnumerable<T>(worksheetIndex, skipCastingErrors, hasHeaderRow).ToList();
        }
    }
}
