using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Core.Extensions
{
    public static class ToExcelExtensions
    {
        /// <summary>
        /// Converts given list of objects to ExcelPackage
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <param name="workSheetName"></param>
        /// <param name="printHeaders"></param>
        /// <param name="tableStyle"></param>
        /// <returns></returns>
        public static ExcelPackage ToExcelPackage<T>(this IList<T> rows, string workSheetName, bool printHeaders = true, TableStyles tableStyle = TableStyles.None)
        {
            var excelFile = new ExcelPackage();
            ExcelWorksheet worksheet = excelFile.Workbook.Worksheets.Add(workSheetName);
            worksheet.Cells["A1"].LoadFromCollection(Collection: rows, PrintHeaders: true, TableStyle: tableStyle);
            excelFile.Save();
            return excelFile;
        }

        /// <summary>
        /// Converts given list of objects to a byte array
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <param name="workSheetName"></param>
        /// <returns></returns>
        public static byte[] ToXlsx<T>(this IList<T> rows, string workSheetName)
        {
            using (ExcelPackage excelPackage = ToExcelPackage(rows, workSheetName))
            {
                return excelPackage.GetAsByteArray();
            }
        }

        /// <summary>
        /// Converts given list of objects to ExcelWorksheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <param name="workSheetName"></param>
        /// <returns></returns>
        public static ExcelWorksheet ToExcelWorksheet<T>(this IList<T> rows, string workSheetName)
        {
            return ToExcelPackage(rows, workSheetName).Workbook.Worksheets.FirstOrDefault(x => x.Name.Equals(workSheetName, StringComparison.InvariantCultureIgnoreCase));
        }
    }
}
