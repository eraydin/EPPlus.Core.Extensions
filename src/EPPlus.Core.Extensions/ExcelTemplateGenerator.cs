using System;
using System.Collections.Generic;
using System.Reflection;

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
            Type type = executingAssembly.FindExcelWorksheetByName(typeName);
                                                                                    
            List<ExcelTableColumnAttributeAndProperyInfo> headerColumns = type.GetExcelTableColumnAttributesWithProperyInfo();

            ExcelWorksheet worksheet = excelPackage.AddWorksheet(type.GetWorksheetName() ?? typeName);

            var rowOffset = 0;

            for (var i = 0; i < headerColumns.Count; i++)
            {
                worksheet.Cells[rowOffset + 1, i + 1].Value = headerColumns[i].ToString();
            }

            return worksheet;
        }
    }
}
