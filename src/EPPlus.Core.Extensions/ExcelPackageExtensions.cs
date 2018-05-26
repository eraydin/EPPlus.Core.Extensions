using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using EPPlus.Core.Extensions.Configuration;

using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.Core.Extensions
{
    public static class ExcelPackageExtensions
    {
        /// <summary>
        ///     Gets all Excel tables in the package
        /// </summary>
        /// <param name="package"></param>
        public static IEnumerable<ExcelTable> GetAllTables(this ExcelPackage package)
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
        ///     Gets an Excel table by name from the package 
        /// </summary>
        /// <param name="package"></param>
        /// <param name="tableName"></param>
        public static ExcelTable GetTable(this ExcelPackage package, string tableName)
        {
            return package.GetAllTables().FirstOrDefault(t => t.Name.Equals(tableName, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        ///     Checks whether a table is existing in the package or not
        /// </summary>
        /// <param name="package"></param>
        /// <param name="tableName"></param>   
        public static bool HasTable(this ExcelPackage package, string tableName)
        {
            return package.GetAllTables().Any(t => t.Name.Equals(tableName, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        ///     Converts the Excel package into a dataset object
        /// </summary>
        /// <param name="package">T</param>
        /// <param name="hasHeaderRow">Indicates whether worksheets have a header row or not.</param>
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
        ///     Converts given worksheet into list of objects as enumerable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="package"></param>
        /// <param name="configurationAction"></param>
        /// <param name="worksheetIndex"></param>
        /// <returns></returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelPackage package, int worksheetIndex = 1, Action<ExcelReadConfiguration<T>> configurationAction = null) where T : class, new() => package.GetWorksheet(worksheetIndex).AsEnumerable(configurationAction);

        /// <summary>
        ///     Converts given worksheet into list of objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="package"></param>
        /// <param name="worksheetIndex"></param>
        /// <param name="configurationAction"></param>
        /// <returns></returns>
        public static List<T> ToList<T>(this ExcelPackage package, int worksheetIndex = 1, Action<ExcelReadConfiguration<T>> configurationAction = null) where T : class, new() => package.AsEnumerable(worksheetIndex, configurationAction).ToList();

        public static ExcelWorksheet AddWorksheet(this ExcelPackage package, string worksheetName) => package.Workbook.Worksheets.Add(worksheetName);

        public static ExcelWorksheet AddWorksheet(this ExcelPackage package, string worksheetName, ExcelWorksheet copyWorksheet) => package.Workbook.Worksheets.Add(worksheetName, copyWorksheet);

        public static ExcelWorksheet GetWorksheet(this ExcelPackage package, string worksheetName) => package.Workbook.GetWorksheet(worksheetName);

        public static ExcelWorksheet GetWorksheet(this ExcelPackage package, int worksheetIndex) => package.Workbook.GetWorksheet(worksheetIndex);
    }
}
