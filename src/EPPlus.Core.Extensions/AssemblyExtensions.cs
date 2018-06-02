using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace EPPlus.Core.Extensions
{
    public static class AssemblyExtensions
    {
        public static List<Type> FindExcelWorksheetTypes(this Assembly executingAssembly)
            => (from type in executingAssembly.GetTypes()
                where type.IsDefined(typeof(ExcelWorksheetAttribute), false)
                select type).ToList();

        public static List<string> GetNamesOfExcelWorksheetTypes(this Assembly executingAssembly)
            => FindExcelWorksheetTypes(executingAssembly).Select(x => x.Name).ToList();

        public static Type FindExcelWorksheetByName(this Assembly executingAssembly, string typeName)
        {
            Type type = FindExcelWorksheetTypes(executingAssembly).FirstOrDefault(x => x.Name.Equals(typeName, StringComparison.InvariantCultureIgnoreCase));

            if (type == null)
            {
                throw new ArgumentNullException(nameof(typeName), $"Type of {typeName} could not found in given assembly.");
            }

            return type;
        }
    }
}
