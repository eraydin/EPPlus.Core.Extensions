using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using EPPlus.Core.Extensions.Attributes;

namespace EPPlus.Core.Extensions
{
    public static class AssemblyExtensions
    {
        public static List<Type> GetTypesMarkedAsExcelWorksheet(this Assembly thisAssembly)
            => (from type in thisAssembly.GetTypes()
                where type.IsDefined(typeof(ExcelWorksheetAttribute), false)
                select type).ToList();

        /// <summary>
        ///     Finds the types which marked as ExcelWorksheet in the assembly, and returns a list of [objectName, worksheetName] pairs
        /// </summary>
        /// <param name="thisAssembly"></param>
        /// <returns>List of [objectName, worksheetName] pairs</returns>
        public static List<KeyValuePair<string, string>> GetExcelWorksheetNamesOfMarkedTypes(this Assembly thisAssembly)
            => GetTypesMarkedAsExcelWorksheet(thisAssembly).Select(x => new KeyValuePair<string, string>(x.Name, x.GetWorksheetName())).ToList();

        public static Type GetExcelWorksheetMarkedTypeByName(this Assembly thisAssembly, string typeName)
            => GetTypesMarkedAsExcelWorksheet(thisAssembly).FirstOrDefault(x => x.Name.Equals(typeName, StringComparison.InvariantCultureIgnoreCase));
    }
}
