using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("EPPlus.Core.Extensions.Tests")]

namespace EPPlus.Core.Extensions
{
    internal static class TypeExtensions
    {
        /// <summary>
        ///     Determines whether given type is nullable or not
        /// </summary>
        /// <param name="type">Type to test</param>
        /// <returns>True if type is nullable</returns>
        public static bool IsNullable(this Type type) => type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);

        /// <summary>
        ///     Tests whether given type is numeric or not
        /// </summary>
        /// <param name="type">Type to test</param>
        /// <returns>True if type is numeric</returns>
        public static bool IsNumeric(this Type type)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        ///     Returns value of the property name
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static object GetPropertyValue(this object obj, string propertyName) => obj.GetType().GetProperty(propertyName)?.GetValue(obj, null);

        /// <summary>
        ///     Returns PropertyInfo and ExcelTableColumnAttribute pairs of given type
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static List<KeyValuePair<PropertyInfo, ExcelTableColumnAttribute>> GetExcelTableColumnAttributes(this Type type)
        {
            List<KeyValuePair<PropertyInfo, ExcelTableColumnAttribute>> propertyAttributePairs = type.GetProperties(BindingFlags.Instance | BindingFlags.Public)
                                                                                                     .Select(property => new KeyValuePair<PropertyInfo, ExcelTableColumnAttribute>(property, property.GetCustomAttributes(typeof(ExcelTableColumnAttribute), true).FirstOrDefault() as ExcelTableColumnAttribute))
                                                                                                     .Where(p => p.Value != null)
                                                                                                     .ToList();

            if (!propertyAttributePairs.Any())
            {
                throw new ArgumentException($"Given object does not have any {nameof(ExcelTableColumnAttribute)}.");
            }

            return propertyAttributePairs;
        }
    }
}
