using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Core.Extensions
{
    public static class ToExcelExtensions
    {
        /// <summary>
        /// Generates an Excel worksheet from a list
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static WorksheetWrapper<T> ToWorksheet<T>(this IList<T> rows, string name, Action<ExcelColumn> configureColumn = null, Action<ExcelRange> configureHeader = null, Action<ExcelRange> configureHeaderRow = null, Action<ExcelRange, T> configureCell = null)
        {
            var worksheet = new WorksheetWrapper<T>()
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
        /// Starts new worksheet on same Excel package
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="K"></typeparam>
        /// <param name="previousSheet"></param>
        /// <param name="rows"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static WorksheetWrapper<T> NextWorksheet<T, K>(this WorksheetWrapper<K> previousSheet, IList<T> rows, string name, Action<ExcelColumn> configureColumn = null, Action<ExcelRange> configureHeader = null, Action<ExcelRange> configureHeaderRow = null, Action<ExcelRange, T> configureCell = null)
        {
            previousSheet.AppendWorksheet();
            var worksheet = new WorksheetWrapper<T>()
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
        /// Adds a column mapping.  If no column mappings are specified all public properties will be used
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="map"></param>
        /// <param name="columnHeader"></param>
        /// <param name="configureColumn"></param>
        /// <param name="configureHeader"></param>
        /// <returns></returns>
        public static WorksheetWrapper<T> WithColumn<T>(this WorksheetWrapper<T> worksheet, Func<T, object> map,
            string columnHeader, Action<ExcelColumn> configureColumn = null, Action<ExcelRange> configureHeader = null, Action<ExcelRange, T> configureCell = null)
        {
            worksheet.Columns.Add(new WorksheetColumn<T>()
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
        /// Adds a title row to the top of the sheet
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

            worksheet.Titles.Add(new WorksheetTitleRow()
            {
                Title = title,
                ConfigureTitle = configureTitle
            });

            return worksheet;
        }

        /// <summary>
        /// Converts given list of objects to ExcelPackage
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <returns></returns>
        public static ExcelPackage ToPackage<T>(this IList<T> rows)
        {
            return rows.ToWorksheet(typeof(T).Name).ToPackage();
        }

        public static ExcelPackage ToPackage<T>(this WorksheetWrapper<T> lastWorksheet)
        {
            lastWorksheet.AppendWorksheet();
            return lastWorksheet.Package;
        }

        public static byte[] ToXlsx<T>(this IList<T> rows)
        {
            return rows.ToWorksheet(typeof(T).Name).ToXlsx();
        }

        public static byte[] ToXlsx<T>(this WorksheetWrapper<T> lastWorksheet)
        {
            lastWorksheet.AppendWorksheet();
            ExcelPackage package = lastWorksheet.Package;

            using (var stream = new MemoryStream())
            {
                package.SaveAs(stream);
                package.Dispose();
                return stream.ToArray();
            }
        }
    }
}
