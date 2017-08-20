using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Core.Extensions
{
    public static class ToExcelExtensions
    {
        public static ExcelPackage ToExcelPackage<T>(this IList<T> rows, string workSheetName, bool printHeaders = true, TableStyles tableStyle = TableStyles.None)
        {
            var excelFile = new ExcelPackage();
            ExcelWorksheet worksheet = excelFile.Workbook.Worksheets.Add(workSheetName);
            worksheet.Cells["A1"].LoadFromCollection(Collection: rows, PrintHeaders: true, TableStyle: tableStyle);
            excelFile.Save();
            return excelFile;
        }

        public static byte[] ToXlsx<T>(this IList<T> rows, string workSheetName)
        {
            using (ExcelPackage excelPackage = ToExcelPackage(rows, workSheetName))
            {
                return excelPackage.GetAsByteArray();
            }
        }

        public static ExcelWorksheet ToExcelWorksheet<T>(this IList<T> rows, string workSheetName)
        {
            return ToExcelPackage(rows, workSheetName).Workbook.Worksheets.FirstOrDefault(x => x.Name.Equals(workSheetName, StringComparison.InvariantCultureIgnoreCase));
        }
    }
}
