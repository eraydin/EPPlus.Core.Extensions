using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;

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
        public static ExcelTable AsExcelTable<T>(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            // Table names should be unique
            string tableName = $"{worksheet.Name}-{new Random(Guid.NewGuid().GetHashCode()).Next(9999)}";
            return worksheet.AsExcelTable<T>(tableName, hasHeaderRow);
        }

        public static ExcelTable AsExcelTable<T>(this ExcelWorksheet worksheet, string tableName, bool hasHeaderRow = true)
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

            worksheet.Tables.Add(worksheet.GetExcelRange(hasHeaderRow), tableName);
            worksheet.Tables[tableName].ShowHeader = false;


            // TODO : 
            PropertyInfo[] propInfo = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);

            // Build property-table column mapping
            for (var i = 0; i < propInfo.Length; i++)
            {
                PropertyInfo property = propInfo[i];
                var mappingAttribute = (ExcelTableColumnAttribute)property.GetCustomAttributes(typeof(ExcelTableColumnAttribute), true).FirstOrDefault();
                if (mappingAttribute != null)
                {
                    if (mappingAttribute.ColumnIndex == 0 && string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                    {
                        worksheet.Tables[tableName].Columns[i].Name = property.Name;
                    }
                    // Column name was specified
                    if (!string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                    {
                        worksheet.Tables[tableName].Columns[i].Name = mappingAttribute.ColumnName;
                    }
                }
            }

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

            IEnumerable<DataColumn> columns = worksheet.AsExcelTable<DataColumn>(!hasHeaderRow).Columns.Select(x => new DataColumn(!hasHeaderRow ? "Column" + x.Id : x.Name));

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
        /// <param name="hasHeaderRow"></param>
        /// <param name="skipCastErrors"></param>
        /// <returns></returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelWorksheet worksheet, bool skipCastErrors = false, bool hasHeaderRow = true) where T : class, new()
        {
            return worksheet.AsExcelTable<T>(hasHeaderRow).AsEnumerable<T>(skipCastErrors);
        }

        /// <summary>
        ///     Returns objects of specified type from the ExcelWorksheet as a list.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="hasHeaderRow"></param>
        /// <param name="skipCastErrors"></param>
        /// <returns></returns>
        public static IList<T> ToList<T>(this ExcelWorksheet worksheet, bool skipCastErrors = false, bool hasHeaderRow = true) where T : class, new()
        {
            return worksheet.AsEnumerable<T>(skipCastErrors, hasHeaderRow).ToList();
        }

        /// <summary>
        ///     Changes value of the specified cell
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static ExcelWorksheet ChangeCellValue(this ExcelWorksheet worksheet, int rowIndex, int columnIndex, object value)
        {
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
            if (!headerTexts.Any())
            {
                return worksheet;
            }

            worksheet.InsertRow(1, 1);

            for (var i = 0; i < headerTexts.Length; i++)
            {
                worksheet.AddLine(1, i + 1, headerTexts[i]);
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
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
            return worksheet.AddLine(rowIndex, 1, values);
        }

        /// <summary>
        ///     Adds a line to the worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddLine(this ExcelWorksheet worksheet, int rowIndex, int startColumnIndex, params object[] values)
        {
            for (var i = 0; i < values.Length; i++)
            {
                worksheet.ChangeCellValue(rowIndex, i + startColumnIndex, values[i]);
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
        /// <returns></returns>
        public static ExcelWorksheet AddObjects<T>(this ExcelWorksheet worksheet, IList<T> items, int startRowIndex, int startColumnIndex=0)
        {
            for (var i = 0; i < items.Count; i++)
            {
                for (int j = startColumnIndex; j < (startColumnIndex + typeof(T).GetProperties().Length); j++)
                {
                    worksheet.AddLine(i + startRowIndex, j + 1, items[i].GetPropertyValue(typeof(T).GetProperties()[j-startColumnIndex].Name));
                }
            }

            return worksheet;
        }

        /// <summary>
        ///      Adds given list of objects to the worksheet with propery selectors
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="items"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="propertySelectors"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddObjects<T>(this ExcelWorksheet worksheet, IList<T> items, int startRowIndex, params Func<T, object>[] propertySelectors)
        {
            return worksheet.AddObjects(items, startRowIndex, 0, propertySelectors);
        }
        
        /// <summary>
        ///     Adds given list of objects to the worksheet with propery selectors
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="items"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="propertySelectors"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddObjects<T>(this ExcelWorksheet worksheet, IList<T> items, int startRowIndex, int startColumnIndex, params Func<T, object>[] propertySelectors)
        {
            if (propertySelectors == null)
            {
                throw new ArgumentException($"{nameof(propertySelectors)} cannot be null");
            }

            for (var i = 0; i < items.Count; i++)
            {
                for (int j = startColumnIndex; j < (startColumnIndex+propertySelectors.Length); j++)
                {
                    worksheet.AddLine(i + startRowIndex, j + 1, propertySelectors[j-startColumnIndex](items[i]));
                }
            }

            return worksheet;
        }
        
        public static ExcelWorksheet SetFont(this ExcelWorksheet worksheet, ExcelAddress address, Font font)
        {
            worksheet.Cells[address.Address].Style.Font.SetFromFont(font);
            return worksheet;
        }

        public static ExcelWorksheet SetFontColor(this ExcelWorksheet worksheet, ExcelAddress address, Color fontColor)
        {
            worksheet.Cells[address.Address].Style.Font.Color.SetColor(fontColor);
            return worksheet;
        }

        public static ExcelWorksheet SetBackgroundColor(this ExcelWorksheet worksheet, ExcelAddress address, Color backgroundColor)
        {
            worksheet.Cells[address.Address].Style.Fill.BackgroundColor.SetColor(backgroundColor);
            return worksheet;
        }

        public static ExcelWorksheet SetHorizontalAlignment(this ExcelWorksheet worksheet, ExcelAddress address, ExcelHorizontalAlignment horizontalAlignment)
        {
            worksheet.Cells[address.Address].Style.HorizontalAlignment = horizontalAlignment;
            return worksheet;
        }

        public static ExcelWorksheet SetVerticalAlignment(this ExcelWorksheet worksheet, ExcelAddress address, ExcelVerticalAlignment verticalAlignment)
        {
            worksheet.Cells[address.Address].Style.VerticalAlignment = verticalAlignment;
            return worksheet;
        }
    }
}
