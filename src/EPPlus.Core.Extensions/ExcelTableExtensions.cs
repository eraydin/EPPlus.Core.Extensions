using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

using EPPlus.Core.Extensions.Configuration;
using EPPlus.Core.Extensions.Exceptions;

using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.Core.Extensions
{  
    public static class ExcelTableExtensions
    {
        /// <summary>
        ///     Returns data bounds of the Excel table with regards to header and totals row visibility
        /// </summary>
        /// <param name="table">Extended object</param>
        /// <returns>Address range</returns>
        public static ExcelAddress GetDataBounds(this ExcelTable table)
        {  
            var dataBounds = new ExcelAddress(
                table.Address.Start.Row + (table.ShowHeader && !table.Address.IsEmptyRange(table.ShowHeader) ? 1 : 0),
                table.Address.Start.Column,
                table.Address.End.Row - (table.ShowTotal ? 1 : 0),
                table.Address.End.Column
                );

            return dataBounds;
        }
        
        /// <summary>
        ///     Validates the Excel table against the generating type.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="table"></param>
        /// <param name="configurationAction"></param>
        /// <returns>An enumerable of <see cref="ExcelExceptionArgs" /> containing</returns>
        public static IEnumerable<ExcelExceptionArgs> Validate<T>(this ExcelTable table, Action<ExcelReadConfiguration<T>> configurationAction = null) where T : class, new()
        {
            ExcelReadConfiguration<T> configuration = DefaultExcelReadConfiguration<T>.Instance;
            configurationAction?.Invoke(configuration);

            List<KeyValuePair<int, PropertyInfo>> mapping = PrepareMappings(table, configuration).ToList();
            var result = new LinkedList<ExcelExceptionArgs>();

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
                        result.AddLast(new ExcelExceptionArgs
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
        ///     Only primitives and enums are supported as property.
        ///     Currently supports only tables with header.
        /// </remarks>
        /// <typeparam name="T">Type to map to. Type should be a class and should have parameterless constructor.</typeparam>
        /// <param name="table">Table object to fetch</param>
        /// <param name="configurationAction"></param>
        /// <returns>An enumerable of the generating type</returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelTable table, Action<ExcelReadConfiguration<T>> configurationAction = null) where T : class, new()
        {
            ExcelReadConfiguration<T> configuration = DefaultExcelReadConfiguration<T>.Instance;
            configurationAction?.Invoke(configuration);

            if (!table.IsEmpty(configuration.HasHeaderRow))
            {
                List<KeyValuePair<int, PropertyInfo>> mapping = PrepareMappings(table, configuration).ToList();

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
                            var exceptionArgs = new ExcelExceptionArgs
                            {
                                ColumnName = table.Columns[map.Key].Name,
                                ExpectedType = property.PropertyType,
                                PropertyName = property.Name,
                                CellValue = cell,
                                CellAddress = new ExcelCellAddress(row, map.Key + table.Address.Start.Column)
                            };

                            if (configuration.ThrowValidationExceptions && ex is ValidationException)
                            {
                                throw new ExcelValidationException(ex.Message, ex)
                                    .WithArguments(exceptionArgs);
                            }

                            if (configuration.ThrowCastingExceptions)
                            {
                                throw new ExcelException(string.Format(configuration.CastingExceptionMessage, exceptionArgs.ColumnName, exceptionArgs.CellAddress.Address, exceptionArgs.CellValue, exceptionArgs.ExpectedType.Name), ex)
                                    .WithArguments(exceptionArgs);
                            }
                        }
                    }

                    configuration.OnCaught?.Invoke(item, row);  
                    yield return item;
                }
            } 
        }  
       
        public static List<T> ToList<T>(this ExcelTable table, Action<ExcelReadConfiguration<T>> configurationAction = null) where T : class, new() => AsEnumerable(table, configurationAction).ToList();

        /// <summary>
        ///     Checks whether given Excel table empty or not
        /// </summary>
        /// <param name="table">Excel table</param>
        /// <param name="hasHeader">'true' as default</param>
        /// <returns>'true' or 'false'</returns>
        public static bool IsEmpty(this ExcelTable table, bool hasHeader = true)
        {
            return table.Address.IsEmptyRange(hasHeader);
        }

        /// <summary>
        ///     Prepares mapping using the type and the attributes decorating its properties
        /// </summary>
        /// <typeparam name="T">Type to parse</typeparam>
        /// <param name="table">Table to get columns from</param>
        /// <param name="configuration"></param>
        /// <returns>A list of mappings from column index to property</returns>
        private static IEnumerable<KeyValuePair<int, PropertyInfo>> PrepareMappings<T>(ExcelTable table, ExcelReadConfiguration<T> configuration)
        {
            // Get only the properties that have ExcelTableColumnAttribute
            List<KeyValuePair<PropertyInfo, ExcelTableColumnAttribute>> propertyAttributePairs = typeof(T).GetExcelTableColumnAttributes();

            // Build property-table column mapping
            foreach (KeyValuePair<PropertyInfo, ExcelTableColumnAttribute> propertyAttributePair in propertyAttributePairs)
            {
                PropertyInfo property = propertyAttributePair.Key;

                ExcelTableColumnAttribute mappingAttribute = propertyAttributePair.Value;

                int col = -1;

                // There is no case when both column name and index is specified since this is excluded by the attribute
                // Neither index, nor column name is specified, use property name
                if (mappingAttribute.ColumnIndex == 0 && string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                {
                    col = table.Columns[property.Name].Position;
                }
                else if (mappingAttribute.ColumnIndex > 0) // Column index was specified
                {
                    col = table.Columns[mappingAttribute.ColumnIndex - 1].Position;
                }
                else if (!string.IsNullOrWhiteSpace(mappingAttribute.ColumnName) && table.Columns.FirstOrDefault(x => x.Name.Equals(mappingAttribute.ColumnName, StringComparison.InvariantCultureIgnoreCase)) != null) // Column name was specified
                {
                    col = table.Columns.First(x => x.Name.Equals(mappingAttribute.ColumnName, StringComparison.InvariantCultureIgnoreCase)).Position;
                }

                if (col == -1)
                {
                    throw new ExcelValidationException(string.Format(configuration.ColumnValidationExceptionMessage, mappingAttribute.ColumnName))
                        .WithArguments(new ExcelExceptionArgs
                                           {
                                               ColumnName = mappingAttribute.ColumnName,
                                               ExpectedType = property.PropertyType,
                                               PropertyName = property.Name,
                                               CellValue = table.WorkSheet.Cells[table.Address.Start.Row, mappingAttribute.ColumnIndex + table.Address.Start.Column].Value,
                                               CellAddress = new ExcelCellAddress(table.Address.Start.Row, mappingAttribute.ColumnIndex + table.Address.Start.Column)
                                           });
                }

                yield return new KeyValuePair<int, PropertyInfo>(col, property);
            }
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
            if (property.PropertyType.IsNullable() && cell != null)
            {
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
            }

            if (type == typeof(DateTime))
            {
                if (!DateTime.TryParse(cell.ToString(), out DateTime parsedDate))
                {
                    parsedDate = DateTime.FromOADate((double)cell);
                }

                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { parsedDate });
            }

            if (type == typeof(bool))
            {
                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { cell });
            }

            if (type.IsEnum)
            {
                if (cell is string) // Support Enum conversion from string...
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
            }

            if (!type.IsEnum && type.IsNumeric())
            {
                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { Convert.ChangeType(cell, type) });
            }

            Validator.ValidateProperty(property.GetValue(item), new ValidationContext(item) { MemberName = property.Name});
        }
    }
}
