using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions
{
    public static class ExcelTemplateGenerator
    {
        public static ExcelPackage GenerateExcelPackage(this Assembly executingAssembly, string typeName)
        {
            var excelPackage = new ExcelPackage();
            excelPackage.GenerateWorksheet(executingAssembly, typeName);
            return excelPackage;
        }  

        public static ExcelWorksheet GenerateWorksheet(this ExcelPackage excelPackage, Assembly executingAssembly, string typeName)
        {
            Type type = executingAssembly.GetTypeByName(typeName);     

            List<KeyValuePair<PropertyInfo, ExcelTableColumnAttribute>> headerColumns = type.GetExcelTableColumnAttributes();

            ExcelWorksheet worksheet = excelPackage.AddWorksheet(typeName);

            var rowOffset = 0;

            for (var i = 0; i < headerColumns.Count; i++)
            {
                string header = !string.IsNullOrEmpty(headerColumns[i].Value.ColumnName) ? headerColumns[i].Value.ColumnName : Regex.Replace(headerColumns[i].Key.Name, "[a-z][A-Z]", m => $"{m.Value[0]} {m.Value[1]}");

                worksheet.Cells[rowOffset + 1, i + 1].Value = header;
            }

            return worksheet;
        }
    }
}
