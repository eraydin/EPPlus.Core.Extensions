using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
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
        /// Export objects to XLSX
        /// </summary>
        /// <typeparam name="T">Type of object</typeparam>
        /// <param name="properties">Class access to the object through its properties</param>
        /// <param name="itemsToExport">The objects to export</param>
        /// <returns></returns>
        public static byte[] ToXlsx<T>(PropertyByName<T>[] properties, IEnumerable<T> itemsToExport)
        {
            using (var stream = new MemoryStream())
            {
                // ok, we can run the real code of the sample now
                using (var xlPackage = new ExcelPackage(stream))
                {
                    // uncomment this line if you want the XML written out to the outputDir
                    //xlPackage.DebugMode = true; 

                    // get handles to the worksheets
                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add(typeof(T).Name);

                    //create Headers and format them 
                    var manager = new PropertyManager<T>(properties.Where(p => !p.Ignore));
                    manager.WriteCaption(worksheet, SetCaptionStyle);

                    var row = 2;
                    foreach (T items in itemsToExport)
                    {
                        manager.CurrentObject = items;
                        manager.WriteToXlsx(worksheet, row++);
                    }

                    xlPackage.Save();
                }
                return stream.ToArray();
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

        private static void SetCaptionStyle(ExcelStyle style)
        {
            style.Fill.PatternType = ExcelFillStyle.Solid;
            style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            style.Font.Bold = true;
        }
    }
}
