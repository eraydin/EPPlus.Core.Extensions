using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;

using EPPlus.Core.Extensions.Configuration;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace EPPlus.Core.Extensions
{
    public static class ExcelWorksheetExtensions
    {
        /// <summary>
        ///     Returns given ExcelWorksheet data bounds as ExcelAddress
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns></returns>
        public static ExcelAddress GetDataBounds(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            return new ExcelAddress(
                worksheet.Dimension.Start.Row + (hasHeaderRow ? 1 : 0),
                worksheet.Dimension.Start.Column,
                worksheet.Dimension.End.Row,
                worksheet.Dimension.End.Column
            );
        }

        /// <summary>
        ///     Returns given ExcelWorksheet data cell ranges as ExcelRange
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns></returns>
        public static ExcelRange GetExcelRange(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            return worksheet.Cells[worksheet.GetDataBounds(hasHeaderRow).Address];
        }

        /// <summary>
        ///     Extracts an ExcelTable from given ExcelWorkSheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns></returns>
        public static ExcelTable AsExcelTable(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            // Table names should be unique
            string tableName = $"{worksheet.Name}-{new Random(Guid.NewGuid().GetHashCode()).Next(9999)}";
            return worksheet.AsExcelTable(tableName, hasHeaderRow);
        }

        public static ExcelTable AsExcelTable(this ExcelWorksheet worksheet, string tableName, bool hasHeaderRow = true)
        {
            if (worksheet.Tables.Any())
            {
                // Has any table on same addresses
                ExcelAddress dataBounds = worksheet.GetDataBounds(false);
                ExcelTable excelTable = worksheet.Tables.FirstOrDefault(x => x.Address.Address.Equals(dataBounds.Address, StringComparison.InvariantCultureIgnoreCase));
                if (excelTable != null)
                {
                    return excelTable;
                }
            }

            worksheet.Tables.Add(worksheet.GetExcelRange(false), tableName);
            worksheet.Tables[tableName].ShowHeader = hasHeaderRow;

            return worksheet.Tables[tableName];
        }

        /// <summary>
        ///     Indicates whether the ExcelWorksheet contains any formula or not
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static bool HasAnyFormula(this ExcelWorksheet worksheet)
        {
            return worksheet.Cells.Any(x => !string.IsNullOrEmpty(x.Formula)) || worksheet.Cells.Any(x => !string.IsNullOrEmpty(x.FormulaR1C1));
        }

        /// <summary>
        ///     Extracts a DataTable from the ExcelWorksheet.
        /// </summary>
        /// <param name="worksheet">The ExcelWorksheet.</param>
        /// <param name="hasHeaderRow">Indicates whether worksheet has a header row or not.</param>
        /// <returns></returns>
        public static DataTable ToDataTable(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            ExcelAddress dataBounds = worksheet.GetDataBounds(hasHeaderRow);

            IEnumerable<DataColumn> columns = worksheet.AsExcelTable(!hasHeaderRow).Columns.Select(x => new DataColumn(!hasHeaderRow ? "Column" + x.Id : x.Name));

            var dataTable = new DataTable(worksheet.Name);
            dataTable.Columns.AddRange(columns.ToArray());

            for (int rowIndex = dataBounds.Start.Row; rowIndex <= dataBounds.End.Row; ++rowIndex)
            {
                ExcelRangeBase[] inputRow = worksheet.Cells[rowIndex, dataBounds.Start.Column, rowIndex, dataBounds.End.Column].ToArray();
                DataRow row = dataTable.Rows.Add();

                for (var j = 0; j < inputRow.Length; ++j)
                {
                    row[j] = inputRow[j].Value;
                }
            }

            return dataTable;
        }

        /// <summary>
        ///     Generic extension method yielding objects of specified type from the ExcelWorksheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="configuration"></param>
        /// <returns></returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelWorksheet worksheet, IExcelConfiguration configuration) where T : class, new()
        {
            return worksheet.AsExcelTable(configuration.HasHeaderRow).AsEnumerable<T>(configuration);
        }

        /// <summary>
        ///     Returns objects of specified type from the ExcelWorksheet as a list.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="configuration"></param>
        /// <returns></returns>
        public static IList<T> ToList<T>(this ExcelWorksheet worksheet, IExcelConfiguration configuration) where T : class, new()
        {
            return worksheet.AsEnumerable<T>(configuration).ToList();
        }

        /// <summary>
        ///     Changes value of the specified cell
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <param name="value"></param>
        /// <param name="configureCell"></param>
        /// <returns></returns>
        public static ExcelWorksheet ChangeCellValue(this ExcelWorksheet worksheet, int rowIndex, int columnIndex, object value, Action<ExcelRange> configureCell = null)
        {
            configureCell?.Invoke(worksheet.Cells[rowIndex, columnIndex]);
            worksheet.Cells[rowIndex, columnIndex].Value = value;
            return worksheet;
        }

        /// <summary>
        ///     Inserts a header line to the top of the Excel worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="headerTexts"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddHeader(this ExcelWorksheet worksheet, params string[] headerTexts)
        {
            return worksheet.AddHeader(null, headerTexts);
        }

        /// <summary>
        ///     Inserts a header line to the top of the Excel worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="configureHeader"></param>
        /// <param name="headerTexts"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddHeader(this ExcelWorksheet worksheet, Action<ExcelRange> configureHeader = null, params string[] headerTexts)
        {
            if (!headerTexts.Any())
            {
                return worksheet;
            }

            worksheet.InsertRow(1, 1);

            for (var i = 0; i < headerTexts.Length; i++)
            {
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                worksheet.AddLine(1, i + 1, configureHeader, headerTexts[i]);
            }

            return worksheet;
        }

        /// <summary>
        ///     Adds a line to the worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddLine(this ExcelWorksheet worksheet, int rowIndex, params object[] values)
        {
            return worksheet.AddLine(rowIndex, 1, null, values);
        }

        /// <summary>
        ///     Adds a line to the worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="configureCells"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddLine(this ExcelWorksheet worksheet, int rowIndex, Action<ExcelRange> configureCells = null, params object[] values)
        {
            return worksheet.AddLine(rowIndex, 1, configureCells, values);
        }

        /// <summary>
        ///     Adds a line to the worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="configureCells"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddLine(this ExcelWorksheet worksheet, int rowIndex, int startColumnIndex, Action<ExcelRange> configureCells = null, params object[] values)
        {
            for (var i = 0; i < values.Length; i++)
            {
                worksheet.ChangeCellValue(rowIndex, i + startColumnIndex, values[i], configureCells);
            }

            return worksheet;
        }

        /// <summary>
        ///     Adds given list of objects to the worksheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="items"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="configureCells"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddObjects<T>(this ExcelWorksheet worksheet, IEnumerable<T> items, int startRowIndex, int startColumnIndex = 1, Action<ExcelRange> configureCells = null)
        {
            for (var i = 0; i < items.Count(); i++)
            {
                for (int j = startColumnIndex; j < startColumnIndex + typeof(T).GetProperties().Length; j++)
                {
                    worksheet.AddLine(i + startRowIndex, j, configureCells, items.ElementAt(i).GetPropertyValue(typeof(T).GetProperties()[j - startColumnIndex].Name));
                }
            }

            return worksheet;
        }

        /// <summary>
        ///     Adds given list of objects to the worksheet with propery selectors
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="items"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="propertySelectors"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddObjects<T>(this ExcelWorksheet worksheet, IEnumerable<T> items, int startRowIndex, params Func<T, object>[] propertySelectors)
        {
            return worksheet.AddObjects(items, startRowIndex, 1, null, propertySelectors);
        }

        /// <summary>
        ///     Adds given list of objects to the worksheet with propery selectors
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="items"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="configureCells"></param>
        /// <param name="propertySelectors"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddObjects<T>(this ExcelWorksheet worksheet, IEnumerable<T> items, int startRowIndex, int startColumnIndex, Action<ExcelRange> configureCells = null, params Func<T, object>[] propertySelectors)
        {
            if (propertySelectors == null)
            {
                throw new ArgumentException($"{nameof(propertySelectors)} cannot be null");
            }

            for (var i = 0; i < items.Count(); i++)
            {
                for (int j = startColumnIndex; j < startColumnIndex + propertySelectors.Length; j++)
                {
                    worksheet.AddLine(i + startRowIndex, j, configureCells, propertySelectors[j - startColumnIndex](items.ElementAt(i)));
                }
            }

            return worksheet;
        }

        /// <summary>
        ///     Returns index and value pairs of columns
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public static IEnumerable<KeyValuePair<int, string>> GetColumns(this ExcelWorksheet worksheet, int rowIndex)
        {
            for (int i = worksheet.Dimension.Start.Column; i <= worksheet.Dimension.End.Column; i++)
            {
                yield return new KeyValuePair<int, string>(i, worksheet.Cells[rowIndex, i, rowIndex, i].Value.ToString());
            }
        }

        /// <summary>
        ///     Checks whether given worksheet address has a value or not
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public static bool CheckIfColumnValueIfNullOrEmpty(this ExcelWorksheet worksheet, int rowIndex, int columnIndex)
        {
            object value = worksheet.Cells[rowIndex, columnIndex, rowIndex, columnIndex].Value;
            return value == null || string.IsNullOrWhiteSpace(value.ToString());
        }

        public static ExcelWorksheet SetFont(this ExcelWorksheet worksheet, Font font)
        {
            return worksheet.SetFont(worksheet.Cells, font);
        }

        public static ExcelWorksheet SetFont(this ExcelWorksheet worksheet, ExcelRange cellRange, Font font)
        {
            worksheet.Cells[cellRange.Address].Style.Font.SetFromFont(font);
            return worksheet;
        }

        public static ExcelWorksheet SetFontColor(this ExcelWorksheet worksheet, Color fontColor)
        {
            return worksheet.SetFontColor(worksheet.Cells, fontColor);
        }

        public static ExcelWorksheet SetFontColor(this ExcelWorksheet worksheet, ExcelRange cellRange, Color fontColor)
        {
            worksheet.Cells[cellRange.Address].Style.Font.Color.SetColor(fontColor);
            return worksheet;
        }

        public static ExcelWorksheet SetBackgroundColor(this ExcelWorksheet worksheet, Color backgroundColor, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            return worksheet.SetBackgroundColor(worksheet.Cells, backgroundColor, fillStyle);
        }

        public static ExcelWorksheet SetBackgroundColor(this ExcelWorksheet worksheet, ExcelRange cellRange, Color backgroundColor, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            worksheet.Cells[cellRange.Address].Style.Fill.PatternType = fillStyle;
            worksheet.Cells[cellRange.Address].Style.Fill.BackgroundColor.SetColor(backgroundColor);
            return worksheet;
        }

        public static ExcelWorksheet SetHorizontalAlignment(this ExcelWorksheet worksheet, ExcelHorizontalAlignment horizontalAlignment)
        {
            return worksheet.SetHorizontalAlignment(worksheet.Cells, horizontalAlignment);
        }

        public static ExcelWorksheet SetHorizontalAlignment(this ExcelWorksheet worksheet, ExcelRange cellRange, ExcelHorizontalAlignment horizontalAlignment)
        {
            worksheet.Cells[cellRange.Address].Style.HorizontalAlignment = horizontalAlignment;
            return worksheet;
        }

        public static ExcelWorksheet SetVerticalAlignment(this ExcelWorksheet worksheet, ExcelVerticalAlignment verticalAlignment)
        {
            return worksheet.SetVerticalAlignment(worksheet.Cells, verticalAlignment);
        }

        public static ExcelWorksheet SetVerticalAlignment(this ExcelWorksheet worksheet, ExcelRange cellRange, ExcelVerticalAlignment verticalAlignment)
        {
            worksheet.Cells[cellRange.Address].Style.VerticalAlignment = verticalAlignment;
            return worksheet;
        }
    }
}
