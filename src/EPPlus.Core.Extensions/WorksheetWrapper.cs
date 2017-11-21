using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.RegularExpressions;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions
{
    public class WorksheetWrapper<T>
    {
        internal string Name { get; set; }

        internal bool AppendHeaderRow { get; set; } = true;

        internal ExcelPackage Package { get; set; }

        internal IEnumerable<T> Rows { get; set; }

        internal IList<WorksheetColumn<T>> Columns { get; set; }

        internal IList<WorksheetTitleRow> Titles { get; set; }

        internal Action<ExcelColumn> ConfigureColumn { get; set; }

        internal Action<ExcelRange> ConfigureHeader { get; set; }

        internal Action<ExcelRange> ConfigureHeaderRow { get; set; }

        internal Action<ExcelRange, T> ConfigureCell { get; set; }

        /// <summary>
        ///     Generates columns for all public properties on the type
        /// </summary>
        /// <returns></returns>
        internal IList<WorksheetColumn<T>> AutoGenerateColumns()
        {
            var columns = new List<WorksheetColumn<T>>();

            List<KeyValuePair<PropertyInfo, ExcelTableColumnAttribute>> propertyAttributePairs = typeof(T).GetExcelTableColumnAttributes<T>();

            foreach (KeyValuePair<PropertyInfo, ExcelTableColumnAttribute> propertyAttributePair in propertyAttributePairs)
            {
                PropertyInfo property = propertyAttributePair.Key;
                ExcelTableColumnAttribute mappingAttribute = propertyAttributePair.Value;

                bool isNullableProperty = property.PropertyType.IsNullable();

                string header = !string.IsNullOrEmpty(mappingAttribute.ColumnName) ? mappingAttribute.ColumnName : Regex.Replace(property.Name, "[a-z][A-Z]", m => $"{m.Value[0]} {m.Value[1]}");

                var column = new WorksheetColumn<T>
                             {
                                 Header = header,
                                 Map = GetGetter<T>(property.Name),
                                 ConfigureColumn = c => c.AutoFit(),
                                 ConfigureHeader = c => { c.Style.Font.Bold = !isNullableProperty; }
                             };
                columns.Add(column);
            }

            return columns;
        }

        /// <summary>
        ///     Wraps creation of an Excel worksheet
        /// </summary>
        internal void AppendWorksheet()
        {
            if (Package == null)
            {
                Package = new ExcelPackage();
            }

            ExcelWorksheet worksheet = Package.Workbook.Worksheets.Add(Name);

            var rowOffset = 0;

            //if no columns specified auto generate them with reflection
            if (Columns == null || !Columns.Any())
            {
                Columns = AutoGenerateColumns();
            }

            //render title rows
            if (Titles != null)
            {
                for (var i = 0; i < Titles.Count; i++)
                {
                    ExcelRange range = worksheet.Cells[rowOffset + 1, 1, rowOffset + 1, Columns.Count];
                    range.Merge = true;
                    range.Value = Titles[i].Title;
                    if (Titles[i].ConfigureTitle != null)
                    {
                        Titles[i].ConfigureTitle(range);
                    }
                }

                rowOffset = rowOffset + Titles.Count;
            }

            //render headers
            if (AppendHeaderRow)
            {
                for (var i = 0; i < Columns.Count; i++)
                {
                    worksheet.Cells[rowOffset + 1, i + 1].Value = Columns[i].Header;

                    if (ConfigureHeader != null)
                    {
                        ConfigureHeader(worksheet.Cells[rowOffset + 1, i + 1]);
                    }

                    if (Columns[i].ConfigureHeader != null)
                    {
                        Columns[i].ConfigureHeader(worksheet.Cells[rowOffset + 1, i + 1]);
                    }
                }

                //configure the header row
                if (ConfigureHeaderRow != null)
                {
                    ConfigureHeaderRow(worksheet.Cells[rowOffset + 1, 1, rowOffset + 1, Columns.Count]);
                }
                else
                {
                    worksheet.Cells[rowOffset + 1, 1, rowOffset + 1, Columns.Count].AutoFilter = true;
                }

                rowOffset++;
            }

            //render data
            if (Rows != null)
            {
                for (var r = 0; r < Rows.Count(); r++)
                {
                    for (var c = 0; c < Columns.Count(); c++)
                    {
                        worksheet.Cells[r + rowOffset + 1, c + 1].Value = Columns[c].Map(Rows.ElementAt(r));

                        ConfigureCell?.Invoke(worksheet.Cells[r + rowOffset + 1, c + 1], Rows.ElementAt(r));
                        if (Columns[c].ConfigureCell != null)
                        {
                            Columns[c].ConfigureCell(worksheet.Cells[r + rowOffset + 1, c + 1], Rows.ElementAt(r));
                        }
                    }
                }
            }

            //configure columns
            for (var i = 0; i < Columns.Count; i++)
            {
                ConfigureColumn?.Invoke(worksheet.Column(i + 1));
                if (Columns[i].ConfigureColumn != null)
                {
                    Columns[i].ConfigureColumn(worksheet.Column(i + 1));
                }
            }
        }

        /// <summary>
        ///     Generates a Func from a propertyName
        /// </summary>
        /// <typeparam name="TP"></typeparam>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private Func<TP, object> GetGetter<TP>(string propertyName)
        {
            ParameterExpression arg = Expression.Parameter(typeof(TP), "x");
            MemberExpression expression = Expression.Property(arg, propertyName);
            UnaryExpression conversion = Expression.Convert(expression, typeof(object));
            return Expression.Lambda<Func<TP, object>>(conversion, arg).Compile();
        }
    }
}
