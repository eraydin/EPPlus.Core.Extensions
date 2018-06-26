using System;
using System.Collections.Generic;
using System.Reflection;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions
{
    public static class ExcelTemplateGenerator
    {
        /// <summary>
        ///     Finds given type name in the assembly, and generates Excel package
        /// </summary>
        /// <param name="executingAssembly"></param>
        /// <param name="typeName"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        public static ExcelPackage GenerateExcelPackage(this Assembly executingAssembly, string typeName, Action<ExcelRange> action = null)
        {
            var excelPackage = new ExcelPackage();
            excelPackage.GenerateWorksheet(executingAssembly, typeName, action);
            return excelPackage;
        }

        /// <summary>
        ///     Finds given type name in the assembly, and generates Excel worksheet
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="executingAssembly"></param>
        /// <param name="typeName"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        public static ExcelWorksheet GenerateWorksheet(this ExcelPackage excelPackage, Assembly executingAssembly, string typeName, Action<ExcelRange> action = null)
        {
            Type type = executingAssembly.GetExcelWorksheetMarkedTypeByName(typeName);

            if (type == null)
            {
                throw new ArgumentNullException(nameof(type), $"Type of {typeName} could not found.");
            }

            List<ExcelTableColumnAttributeAndPropertyInfo> headerColumns = type.GetExcelTableColumnAttributesWithProperyInfo();

            ExcelWorksheet worksheet = excelPackage.AddWorksheet(type.GetWorksheetName() ?? typeName);

            for (var i = 0; i < headerColumns.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = headerColumns[i].ToString();
                action?.Invoke(worksheet.Cells[1, i + 1]);
            }

            return worksheet;
        }
    }
}
