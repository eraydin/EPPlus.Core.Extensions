using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using JetBrains.Annotations;

namespace EPPlus.Core.Extensions
{
    public static class AssemblyExtensions
    {
        public static List<Type> GetExcelWorksheetMarkedTypes(this Assembly thisAssembly)
            => (from type in thisAssembly.GetTypes()
                where type.IsDefined(typeof(ExcelWorksheetAttribute), false)
                select type).ToList();

        /// <summary>
        ///     Finds ExcelWorksheet marked types in given assembly, and returns a list of [objectName, worksheetName] pairs
        /// </summary>
        /// <param name="thisAssembly"></param>
        /// <returns>List of [objectName, worksheetName] pairs</returns>
        public static List<KeyValuePair<string, string>> GetExcelWorksheetNamesOfMarkedTypes(this Assembly thisAssembly)
            => GetExcelWorksheetMarkedTypes(thisAssembly).Select(x => new KeyValuePair<string, string>(x.Name, x.GetWorksheetName())).ToList();

        [CanBeNull]
        public static Type GetExcelWorksheetMarkedTypeByName(this Assembly thisAssembly, string typeName)
            => GetExcelWorksheetMarkedTypes(thisAssembly).FirstOrDefault(x => x.Name.Equals(typeName, StringComparison.InvariantCultureIgnoreCase));
    }
}
