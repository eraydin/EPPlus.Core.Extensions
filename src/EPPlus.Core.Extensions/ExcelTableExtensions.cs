using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using EPPlus.Core.Extensions.Configuration;
using EPPlus.Core.Extensions.Validation;

using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.Core.Extensions
{
    /// <summary>
    ///     Class holds extensions on ExcelTable object
    /// </summary>
    public static class ExcelTableExtensions
    {
        /// <summary>
        ///     Returns table data bounds with regards to header and totals row visibility
        /// </summary>
        /// <param name="table">Extended object</param>
        /// <returns>Address range</returns>
        public static ExcelAddress GetDataBounds(this ExcelTable table)
        {
            return new ExcelAddress(
                table.Address.Start.Row + (table.ShowHeader ? 1 : 0),
                table.Address.Start.Column,
                table.Address.End.Row - (table.ShowTotal ? 1 : 0),
                table.Address.End.Column
            );
        }

        /// <summary>
        ///     Validates the excel table against the generating type.
        /// </summary>
        /// <typeparam name="T">Generating class type</typeparam>
        /// <param name="table">Extended object</param>
        /// <returns>An enumerable of <see cref="ExcelTableExceptionArgs" /> containing </returns>
        public static IEnumerable<ExcelTableExceptionArgs> Validate<T>(this ExcelTable table) where T : class, new()
        {
            IList mapping = PrepareMappings<T>(table);
            var result = new LinkedList<ExcelTableExceptionArgs>();

            ExcelAddress bounds = table.GetDataBounds();

            var item = (T)Activator.CreateInstance(typeof(T));

            // Parse table
            for (int row = bounds.Start.Row; row <= bounds.End.Row; row++)
            {
                foreach (KeyValuePair<int, PropertyInfo> map in mapping)
                {
                    object cell = table.WorkSheet.Cells[row, map.Key + table.Address.Start.Column].Value;

                    PropertyInfo property = map.Value;

                    try
                    {
                        TrySetProperty(item, property, cell);
                    }
                    catch
                    {
                        result.AddLast(new ExcelTableExceptionArgs
                        {
                            ColumnName = table.Columns[map.Key].Name,
                            ExpectedType = property.PropertyType,
                            PropertyName = property.Name,
                            CellValue = cell,
                            CellAddress = new ExcelCellAddress(row, map.Key + table.Address.Start.Column)
                        });
                    }
                }
            }

            return result;
        }

        /// <summary>
        ///     Generic extension method yielding objects of specified type from table.
        /// </summary>
        /// <remarks>
        ///     Exceptions are not catched. It works on all or nothing basis.
        ///     Only primitives and enums are supported as property.
        ///     Currently supports only tables with header.
        /// </remarks>
        /// <typeparam name="T">Type to map to. Type should be a class and should have parameterless constructor.</typeparam>
        /// <param name="table">Table object to fetch</param>
        /// <param name="configurationAction"></param>
        /// <returns>An enumerable of the generating type</returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelTable table, Action<IExcelConfiguration> configurationAction = null) where T : class, new()
        {
            IExcelConfiguration configuration = new DefaultExcelConfiguration();
            configurationAction?.Invoke(configuration);

            IList mapping = PrepareMappings<T>(table);

            ExcelAddress bounds = table.GetDataBounds();

            // Parse table
            for (int row = bounds.Start.Row; row <= bounds.End.Row; row++)
            {
                var item = (T)Activator.CreateInstance(typeof(T));

                foreach (KeyValuePair<int, PropertyInfo> map in mapping)
                {
                    object cell = table.WorkSheet.Cells[row, map.Key + table.Address.Start.Column].Value;

                    PropertyInfo property = map.Value;

                    try
                    {
                        TrySetProperty(item, property, cell);
                    }
                    catch (Exception ex)
                    {
                        if (!configuration.SkipCastingErrors)
                        {
                            var exceptionArgs = new ExcelTableExceptionArgs
                            {
                                ColumnName = table.Columns[map.Key].Name,
                                ExpectedType = property.PropertyType,
                                PropertyName = property.Name,
                                CellValue = cell,
                                CellAddress = new ExcelCellAddress(row, map.Key + table.Address.Start.Column)
                            };

                            throw new ExcelTableConvertException($"The expected type of '{exceptionArgs.PropertyName}' property is '{exceptionArgs.ExpectedType.Name}', but the cell [{exceptionArgs.CellAddress.Address}] contains an invalid value.",
                                ex, exceptionArgs
                            );
                        }
                    }
                }

                // TODO:
                if (!configuration.SkipValidationErrors)
                {
                    // Validate parsed object according to data annotations
                    item.Validate(row);
                }

                yield return item;
            }
        }

        /// <summary>
        ///     Returns objects of specified type from table as list.
        /// </summary>
        /// <remarks>
        ///     Exceptions are not catched. It works on all or nothing basis.
        ///     Only primitives and enums are supported as property.
        ///     Currently supports only tables with header.
        /// </remarks>
        /// <typeparam name="T">Type to map to. Type should be a class and should have parameterless constructor.</typeparam>
        /// <param name="table">Table object to fetch</param>
        /// <param name="configurationAction"></param>
        /// <returns>An enumerable of the generating type</returns>
        public static IList<T> ToList<T>(this ExcelTable table, Action<IExcelConfiguration> configurationAction = null) where T : class, new()
        {
            return AsEnumerable<T>(table, configurationAction).ToList();
        }

        /// <summary>
        ///     Prepares mapping using the type and the attributes decorating its properties
        /// </summary>
        /// <typeparam name="T">Type to parse</typeparam>
        /// <param name="table">Table to get columns from</param>
        /// <returns>A list of mappings from column index to property</returns>
        private static IList PrepareMappings<T>(ExcelTable table)
        {
            IList mapping = new List<KeyValuePair<int, PropertyInfo>>();

            PropertyInfo[] propInfo = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);

            // Build property-table column mapping
            foreach (PropertyInfo property in propInfo)
            {
                var mappingAttribute = (ExcelTableColumnAttribute)property.GetCustomAttributes(typeof(ExcelTableColumnAttribute), true).FirstOrDefault();
                if (mappingAttribute != null)
                {
                    int col = -1;

                    // There is no case when both column name and index is specified since this is excluded by the attribute
                    // Neither index, nor column name is specified, use property name
                    if (mappingAttribute.ColumnIndex == 0 && string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                    {
                        col = table.Columns[property.Name].Position;
                    }

                    // Column index was specified
                    if (mappingAttribute.ColumnIndex > 0)
                    {
                        col = table.Columns[mappingAttribute.ColumnIndex - 1].Position;
                    }

                    // Column name was specified
                    if (!string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                    {
                        if (table.Columns.First(x => x.Name.Equals(mappingAttribute.ColumnName, StringComparison.InvariantCultureIgnoreCase)) != null)
                        {
                            col = table.Columns.First(x => x.Name.Equals(mappingAttribute.ColumnName, StringComparison.InvariantCultureIgnoreCase)).Position;
                        }
                    }

                    if (col == -1)
                    {
                        throw new ArgumentException($"{mappingAttribute.ColumnName} column could not found on the worksheet");
                    }

                    mapping.Add(new KeyValuePair<int, PropertyInfo>(col, property));
                }
            }

            return mapping;
        }

        /// <summary>
        ///     Tries to set property of item
        /// </summary>
        /// <param name="item">target object</param>
        /// <param name="property">property to be set</param>
        /// <param name="cell">cell value</param>
        private static void TrySetProperty(object item, PropertyInfo property, object cell)
        {
            Type type = property.PropertyType;
            Type itemType = item.GetType();

            // If type is nullable, get base type instead
            if (property.PropertyType.IsNullable())
            {
                if (cell == null)
                {
                    return; // If it is nullable, and we have null we should not waste time
                }

                type = type.GetGenericArguments()[0];
            }

            if (type == typeof(string))
            {
                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { cell?.ToString() });

                return;
            }

            if (type == typeof(DateTime))
            {
                DateTime d = DateTime.Parse(cell.ToString());

                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { d });

                return;
            }

            if (type == typeof(bool))
            {
                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { cell });

                return;
            }

            if (type.IsEnum)
            {
                if (cell.GetType() == typeof(string)) // Support Enum conversion from string...
                {
                    itemType.InvokeMember(
                        property.Name,
                        BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                        null,
                        item,
                        new object[] { Enum.Parse(type, cell.ToString(), true) });
                }
                else // ...and numeric cell value
                {
                    Type underType = type.GetEnumUnderlyingType();

                    itemType.InvokeMember(
                        property.Name,
                        BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                        null,
                        item,
                        new object[] { Enum.ToObject(type, Convert.ChangeType(cell, underType)) });
                }

                return;
            }

            if (type.IsNumeric())
            {
                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { Convert.ChangeType(cell, type) });
            }
        }
    }
}
