using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using EPPlus.Core.Extensions.Configuration;

using OfficeOpenXml;
using OfficeOpenXml.Table;

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
        /// <param name="package">The ExcelPackage object</param>
        /// <returns>Enumeration of ExcelTables</returns>
        public static IEnumerable<ExcelTable> GetTables(this ExcelPackage package)
        {
            foreach (ExcelWorksheet ws in package.Workbook.Worksheets)
            {
                foreach (ExcelTable t in ws.Tables)
                {
                    yield return t;
                }
            }
        }

        /// <summary>
        ///     Returns concrete ExcelTable by its name
        /// </summary>
        /// <param name="package">The ExcelPackage object</param>
        /// <param name="name">Name of the table</param>
        /// <returns>ExcelTable object if found, null if not</returns>
        public static ExcelTable GetTable(this ExcelPackage package, string name)
        {
            return package.GetTables().FirstOrDefault(t => t.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        ///     Checks that given table name is in the ExcelPackage or not
        /// </summary>
        /// <param name="package">The ExcelPackage object</param>
        /// <param name="name">Name of the table</param>
        /// <returns>Result of search as bool</returns>
        public static bool HasTable(this ExcelPackage package, string name)
        {
            return package.GetTables().Any(t => t.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        ///     Extracts a DataSet from the ExcelPackage.
        /// </summary>
        /// <param name="package">The ExcelPackage.</param>
        /// <param name="hasHeaderRow">Indicates whether worksheet has a header row or not.</param>
        /// <returns></returns>
        public static DataSet ToDataSet(this ExcelPackage package, bool hasHeaderRow = true)
        {
            var dataSet = new DataSet();

            foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
            {
                dataSet.Tables.Add(worksheet.ToDataTable(hasHeaderRow));
            }

            return dataSet;
        }

        /// <summary>
        ///     Yields objects of specified type from given ExcelPackage
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="package"></param>
        /// <param name="configurationAction"></param>
        /// <param name="worksheetIndex"></param>
        /// <returns></returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelPackage package, int worksheetIndex = 1, Action<IExcelReadConfiguration<T>> configurationAction = null) where T : class, new()
        {
            return package.Workbook.Worksheets[worksheetIndex].AsEnumerable(configurationAction);
        }

        /// <summary>
        ///     Converts given ExcelPackage to list of objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="package"></param>
        /// <param name="worksheetIndex"></param>
        /// <param name="configurationAction"></param>
        /// <returns></returns>
        public static List<T> ToList<T>(this ExcelPackage package, int worksheetIndex = 1, Action<IExcelReadConfiguration<T>> configurationAction = null) where T : class, new()
        {
            return package.AsEnumerable(worksheetIndex, configurationAction).ToList();
        }
    }
}
