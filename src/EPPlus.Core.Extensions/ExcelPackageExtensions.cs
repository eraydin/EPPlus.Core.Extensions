using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace EPPlus.Core.Extensions
{
    /// <summary>
    /// Class holds extensions on ExcelPackage object
    /// </summary>
    public static class ExcelPackageExtensions
    {
        /// <summary>
        /// Returns all table names in the opened worksheet
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
        /// Returns concrete ExcelTable by its name 
        /// </summary>
        /// <param name="excelPackage">The ExcelPackage object</param>
        /// <param name="name">Name of the table</param>
        /// <returns>ExcelTable object if found, null if not</returns>
        public static ExcelTable GetTable(this ExcelPackage excelPackage, string name)
        {
            return excelPackage.GetTables().FirstOrDefault(t => t.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Checks that given table name is in the ExcelPackage or not
        /// </summary>
        /// <param name="excel">The ExcelPackage object</param>
        /// <param name="name">Name of the table</param>
        /// <returns>Result of search as bool</returns>
        public static bool HasTable(this ExcelPackage excel, string name)
        {
            return excel.GetTables().Any(t => t.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Extracts a DataSet from the ExcelPackage.
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
        ///     Creates a new instance of the ExcelPackage class based on a byte array
        /// </summary>
        /// <param name="buffer">The byte array</param>
        /// <returns>An ExcelPackages</returns>
        public static ExcelPackage ToExcelPackage(this byte[] buffer)
        {
            using (var memoryStream = new MemoryStream(buffer))
            {
                return new ExcelPackage(memoryStream);
            }
        }

        /// <summary>
        ///     Creates a new instance of the ExcelPackage class based on a byte array
        /// </summary>
        /// <param name="buffer">The byte array</param>
        /// <param name="password">The password to decrypt the document</param>
        /// <returns>An ExcelPackages</returns>
        public static ExcelPackage ToExcelPackage(this byte[] buffer, string password)
        {
            if (!string.IsNullOrEmpty(password))
            {
                using (var memoryStream = new MemoryStream(buffer))
                {
                    return new ExcelPackage(memoryStream, password);
                }
            }

            return ToExcelPackage(buffer);
        }
    }
}
