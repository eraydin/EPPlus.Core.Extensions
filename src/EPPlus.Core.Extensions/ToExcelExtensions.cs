using System;
using System.Collections.Generic;
using System.IO;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions
{
    public static class ToExcelExtensions
    {
        /// <summary>
        ///     Generates an Excel worksheet from a list
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <param name="name"></param>
        /// <param name="configureColumn"></param>
        /// <param name="configureHeader"></param>
        /// <param name="configureHeaderRow"></param>
        /// <param name="configureCell"></param>
        /// <returns></returns>
        public static WorksheetWrapper<T> ToWorksheet<T>(this IEnumerable<T> rows, string name, Action<ExcelColumn> configureColumn = null, Action<ExcelRange> configureHeader = null, Action<ExcelRange> configureHeaderRow = null, Action<ExcelRange, T> configureCell = null)
        {
            var worksheet = new WorksheetWrapper<T>
            {
                Name = name,
                Package = new ExcelPackage(),
                Rows = rows,
                Columns = new List<WorksheetColumn<T>>(),
                ConfigureHeader = configureHeader,
                ConfigureColumn = configureColumn,
                ConfigureHeaderRow = configureHeaderRow,
                ConfigureCell = configureCell
            };
            return worksheet;
        }

        /// <summary>
        ///     Starts new worksheet on same Excel package
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="K"></typeparam>
        /// <param name="previousSheet"></param>
        /// <param name="rows"></param>
        /// <param name="name"></param>
        /// <param name="configureColumn"></param>
        /// <param name="configureHeader"></param>
        /// <param name="configureHeaderRow"></param>
        /// <param name="configureCell"></param>
        /// <returns></returns>
        public static WorksheetWrapper<T> NextWorksheet<T, K>(this WorksheetWrapper<K> previousSheet, IEnumerable<T> rows, string name, Action<ExcelColumn> configureColumn = null, Action<ExcelRange> configureHeader = null, Action<ExcelRange> configureHeaderRow = null, Action<ExcelRange, T> configureCell = null)
        {
            previousSheet.AppendWorksheet();
            var worksheet = new WorksheetWrapper<T>
            {
                Name = name,
                Package = previousSheet.Package,
                Rows = rows,
                Columns = new List<WorksheetColumn<T>>(),
                ConfigureHeader = configureHeader ?? previousSheet.ConfigureHeader,
                ConfigureColumn = configureColumn ?? previousSheet.ConfigureColumn,
                ConfigureHeaderRow = configureHeaderRow ?? previousSheet.ConfigureHeaderRow,
                ConfigureCell = configureCell
            };
            return worksheet;
        }

        /// <summary>
        ///     Adds a column mapping.  If no column mappings are specified all public properties will be used
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="map"></param>
        /// <param name="columnHeader"></param>
        /// <param name="configureColumn"></param>
        /// <param name="configureHeader"></param>
        /// <param name="configureCell"></param>
        /// <returns></returns>
        public static WorksheetWrapper<T> WithColumn<T>(this WorksheetWrapper<T> worksheet, Func<T, object> map,
            string columnHeader, Action<ExcelColumn> configureColumn = null, Action<ExcelRange> configureHeader = null, Action<ExcelRange, T> configureCell = null)
        {
            worksheet.Columns.Add(new WorksheetColumn<T>
            {
                Map = map,
                ConfigureHeader = configureHeader,
                ConfigureColumn = configureColumn,
                Header = columnHeader,
                ConfigureCell = configureCell
            });
            return worksheet;
        }

        /// <summary>
        ///     Adds a title row to the top of the sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="title"></param>
        /// <param name="configureTitle"></param>
        /// <returns></returns>
        public static WorksheetWrapper<T> WithTitle<T>(this WorksheetWrapper<T> worksheet, string title, Action<ExcelRange> configureTitle = null)
        {
            if (worksheet.Titles == null)
            {
                worksheet.Titles = new List<WorksheetTitleRow>();
            }

            worksheet.Titles.Add(new WorksheetTitleRow
            {
                Title = title,
                ConfigureTitle = configureTitle
            });

            return worksheet;
        }

        /// <summary>
        ///     Indicates that the worksheet should not contain a header row
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static WorksheetWrapper<T> WithoutHeader<T>(this WorksheetWrapper<T> worksheet)
        {
            worksheet.AppendHeaderRow = false;
            return worksheet;
        }

        /// <summary>
        ///     Converts given list of objects to ExcelPackage
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <param name="addHeaderRow"></param>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        public static ExcelPackage ToExcelPackage<T>(this IEnumerable<T> rows, bool addHeaderRow = true, string worksheetName = null)
        {
            WorksheetWrapper<T> worksheet = rows.ToWorksheet(string.IsNullOrEmpty(worksheetName) ? typeof(T).Name : worksheetName);

            if (!addHeaderRow)
            {
                worksheet.WithoutHeader();
            }

            return worksheet.ToExcelPackage();
        }

        public static ExcelPackage ToExcelPackage<T>(this WorksheetWrapper<T> lastWorksheet)
        {
            lastWorksheet.AppendWorksheet();
            return lastWorksheet.Package;
        }

        /// <summary>
        ///     Creates a new instance of the ExcelPackage class based on a byte array
        /// </summary>
        /// <param name="buffer">The byte array</param>
        /// <returns>An ExcelPackages</returns>
        public static ExcelPackage GetAsExcelPackage(this byte[] buffer)
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
        public static ExcelPackage GetAsExcelPackage(this byte[] buffer, string password)
        {
            if (!string.IsNullOrEmpty(password))
            {
                using (var memoryStream = new MemoryStream(buffer))
                {
                    return new ExcelPackage(memoryStream, password);
                }
            }

            return buffer.ToExcelPackage();
        }

        /// <summary>
        ///     Converts list of items to Excel and returns the Excel file as a bytearray.
        /// </summary>
        /// <typeparam name="T">Type of object</typeparam>
        /// <param name="rows">List of objects</param>
        /// <param name="addHeaderRow">Add header row to worksheet</param>
        /// <returns></returns>
        public static byte[] ToXlsx<T>(this IEnumerable<T> rows, bool addHeaderRow = true)
        {
            WorksheetWrapper<T> worksheet = rows.ToWorksheet(typeof(T).Name);

            if (!addHeaderRow)
            {
                worksheet.WithoutHeader();
            }

            return worksheet.ToXlsx();
        }

        /// <summary>
        ///     Returns the Excel file as a bytearray.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="lastWorksheet"></param>
        /// <returns></returns>
        public static byte[] ToXlsx<T>(this WorksheetWrapper<T> lastWorksheet)
        {
            lastWorksheet.AppendWorksheet();

            using (var stream = new MemoryStream())
            {
                using (ExcelPackage package = lastWorksheet.Package)
                {
                    package.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }
    }
}
