using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

using EPPlus.Core.Extensions.Configuration;
using EPPlus.Core.Extensions.Exceptions;

using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.Core.Extensions
{
    public static class ExcelWorksheetExtensions
    {
        /// <summary>
        ///     Returns data bounds of the worksheet as ExcelAddress
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns></returns>
        public static ExcelAddress GetDataBounds(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            ExcelAddressBase valuedDimension = worksheet.GetValuedDimension() ?? worksheet.Dimension;

            if (valuedDimension == null)
            {
                return null;
            }

            return new ExcelAddress(
                valuedDimension.Start.Row + (hasHeaderRow && valuedDimension.Start.Row != valuedDimension.End.Row ? 1 : 0),
                valuedDimension.Start.Column,
                valuedDimension.End.Row,
                valuedDimension.End.Column
                );
        }

        /// <summary>
        ///     Returns cell ranges of the worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns></returns>
        public static ExcelRange GetExcelRange(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            ExcelAddress dataBounds = worksheet.GetDataBounds(hasHeaderRow);

            if (dataBounds == null)
            {
                return null;
            }

            return worksheet.Cells[dataBounds.Address];
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
            string tableName = $"Table{new Random(Guid.NewGuid().GetHashCode()).Next(99999)}";
            return worksheet.AsExcelTable(tableName, hasHeaderRow);
        }

        public static ExcelTable GetTable(this ExcelWorksheet worksheet, string tableName) => worksheet.Tables.FirstOrDefault(x => x.Name == tableName);

        public static ExcelTable GetTable(this ExcelWorksheet worksheet, int tableIndex) => worksheet.Tables[tableIndex];

        /// <summary>
        ///     Converts given worksheet into ExcelTable
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="tableName"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns></returns>
        public static ExcelTable AsExcelTable(this ExcelWorksheet worksheet, string tableName, bool hasHeaderRow)
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

            ExcelRange dataRange = worksheet.GetExcelRange(false) ?? worksheet.Cells[1, 1, 1, 1];

            worksheet.Tables.Add(dataRange, tableName);
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

            var dataTable = new DataTable(worksheet.Name);

            if (dataBounds == null)
            {
                return dataTable;
            }

            IEnumerable<DataColumn> columns = worksheet.AsExcelTable(hasHeaderRow).Columns.Select(x => new DataColumn(!hasHeaderRow ? "Column" + x.Id : x.Name));

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
        ///     Converts the worksheet into list of objects as enumerable
        /// </summary>
        /// <typeparam name="T">Type of object</typeparam>
        /// <param name="worksheet"></param>
        /// <param name="configurationAction"></param>
        /// <returns></returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelWorksheet worksheet, Action<ExcelReadConfiguration<T>> configurationAction = null) where T : class, new()
        {
            ExcelReadConfiguration<T> configuration = DefaultExcelReadConfiguration<T>.Instance;
            configurationAction?.Invoke(configuration);

            return worksheet.AsExcelTable(configuration.HasHeaderRow).AsEnumerable(configurationAction);
        }

        /// <summary>
        ///     Converts the worksheet into list of objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="configurationAction"></param>
        /// <returns></returns>
        public static List<T> ToList<T>(this ExcelWorksheet worksheet, Action<ExcelReadConfiguration<T>> configurationAction = null) where T : class, new() => worksheet.AsEnumerable(configurationAction).ToList();

        public static ExcelWorksheet ChangeCellValue(this ExcelWorksheet worksheet, int rowIndex, int columnIndex, object value, Action<ExcelRange> configureCell = null)
        {
            configureCell?.Invoke(worksheet.Cells[rowIndex, columnIndex]);
            worksheet.Cells[rowIndex, columnIndex].Value = value;
            return worksheet;
        }

        /// <summary>
        ///     Inserts a header line to the top of the worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="headerTexts"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddHeader(this ExcelWorksheet worksheet, params string[] headerTexts) => worksheet.AddHeader(null, headerTexts);

        /// <summary>
        ///     Inserts a header line to the top of the worksheet
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
        public static ExcelWorksheet AddLine(this ExcelWorksheet worksheet, int rowIndex, params object[] values) => worksheet.AddLine(rowIndex, 1, null, values);

        /// <summary>
        ///     Adds a line to the worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="configureCells"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddLine(this ExcelWorksheet worksheet, int rowIndex, Action<ExcelRange> configureCells = null, params object[] values) => worksheet.AddLine(rowIndex, 1, configureCells, values);

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
        public static ExcelWorksheet AddObjects<T>(this ExcelWorksheet worksheet, IEnumerable<T> items, int startRowIndex, params Func<T, object>[] propertySelectors) => worksheet.AddObjects(items, startRowIndex, 1, null, propertySelectors);

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
            ExcelAddressBase valuedDimension = worksheet.GetValuedDimension();

            if (valuedDimension != null)
            {
                for (int i = valuedDimension.Start.Column; i <= valuedDimension.End.Column; i++)
                {
                    yield return new KeyValuePair<int, string>(i, worksheet.Cells[rowIndex, i, rowIndex, i].Text);
                }
            }
        }

        /// <summary>
        ///     Deletes a column from worksheet by using column header text
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="headerText"></param>
        /// <returns></returns>
        public static ExcelWorksheet DeleteColumn(this ExcelWorksheet worksheet, string headerText)
        {
            ExcelAddressBase valuedDimension = worksheet.GetValuedDimension();

            ExcelRangeBase headerColumn = worksheet.Cells[valuedDimension.Start.Row, valuedDimension.Start.Column, valuedDimension.Start.Row, valuedDimension.End.Column].FirstOrDefault(x => x.Text.Equals(headerText, StringComparison.InvariantCultureIgnoreCase));

            if (headerColumn != null)
            {
                worksheet.DeleteColumn(headerColumn.Start.Column);
            }

            return worksheet;
        }

        /// <summary>
        ///     Deletes columns from worksheet by using given header text
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="headerText"></param>
        /// <returns></returns>
        public static ExcelWorksheet DeleteColumns(this ExcelWorksheet worksheet, string headerText)
        {
            ExcelAddressBase valuedDimension = worksheet.GetValuedDimension();

            int count = worksheet.Cells[valuedDimension.Start.Row, valuedDimension.Start.Column, valuedDimension.Start.Row, valuedDimension.End.Column].Count(x => x.Text.Equals(headerText, StringComparison.InvariantCultureIgnoreCase));

            for (var i = 0; i < count; i++)
            {
                worksheet.DeleteColumn(headerText);
            }

            return worksheet;
        }

        /// <summary>
        ///     Checks and throws the <see cref="ExcelValidationException"/> if column value is wrong on specified index
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <param name="expectedValue"></param>
        /// <param name="exceptionMessage">Custom exception message with format parameters: columnIndex, expectedValue</param>
        public static void CheckAndThrowColumn(this ExcelWorksheet worksheet, int rowIndex, int columnIndex, string expectedValue, string exceptionMessage = null)
        {
            if (!worksheet.GetColumns(rowIndex).Any(x => x.Value == expectedValue && x.Key == columnIndex))
            {
                if (!string.IsNullOrEmpty(exceptionMessage))
                {
                    throw new ExcelValidationException(string.Format(exceptionMessage, columnIndex, expectedValue));
                }

                throw new ExcelValidationException($"The {columnIndex}. column of worksheet should be '{expectedValue}'.");
            }
        }

        /// <summary>
        ///     Checks and throws the <see cref="ExcelValidationException"/> if the worksheet has any formula
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="withMessage"></param>
        public static void CheckAndThrowIfThereIsAnyFormula(this ExcelWorksheet sheet, string withMessage)
        {
            if (sheet.HasAnyFormula())
            {
                throw new ExcelValidationException(withMessage);
            }
        }

        /// <summary>
        ///     Checks and throws the <see cref="ExcelValidationException"/> if header columns does not match with properties of object 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="headerRowIndex"></param>
        /// <param name="formattedExceptionMessage"></param>
        public static void CheckHeadersAndThrow<T>(this ExcelWorksheet worksheet, int headerRowIndex, string formattedExceptionMessage = null)
        {
            List<KeyValuePair<PropertyInfo, ExcelTableColumnAttribute>> propertyAttributePairs = typeof(T).GetExcelTableColumnAttributes();

            for (var i = 0; i < propertyAttributePairs.Count; i++)
            {
                string columnName = !string.IsNullOrEmpty(propertyAttributePairs[i].Value.ColumnName) ? propertyAttributePairs[i].Value.ColumnName : propertyAttributePairs[i].Key.Name;
                worksheet.CheckAndThrowColumn(headerRowIndex, i + 1, columnName, formattedExceptionMessage);
            }
        }

        /// <summary>
        ///     Checks whether given worksheet address has a value or not
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public static bool CheckColumnValueIsNullOrEmpty(this ExcelWorksheet worksheet, int rowIndex, int columnIndex)
        {
            object value = worksheet.Cells[rowIndex, columnIndex, rowIndex, columnIndex].Value;
            return string.IsNullOrWhiteSpace(value?.ToString());
        }

        /// <summary>
        ///     Gets valued dimensions of worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static ExcelAddressBase GetValuedDimension(this ExcelWorksheet worksheet)
        {
            ExcelAddressBase dimension = worksheet.Dimension;

            if (dimension == null)
            {
                return null;
            }

            ExcelRange cells = worksheet.Cells[dimension.Address];
            int minRow = 0, minCol = 0, maxRow = 0, maxCol = 0;
            var hasValue = false;
            foreach (ExcelRangeBase cell in cells.Where(cell => cell.Value != null))
            {
                if (!hasValue)
                {
                    minRow = cell.Start.Row;
                    minCol = cell.Start.Column;
                    maxRow = cell.End.Row;
                    maxCol = cell.End.Column;
                    hasValue = true;
                }
                else
                {
                    if (cell.Start.Column < minCol)
                    {
                        minCol = cell.Start.Column;
                    }
                    if (cell.End.Row > maxRow)
                    {
                        maxRow = cell.End.Row;
                    }
                    if (cell.End.Column > maxCol)
                    {
                        maxCol = cell.End.Column;
                    }
                }
            }

            return hasValue ? new ExcelAddressBase(minRow, minCol, maxRow, maxCol) : null;
        }
    }
}
