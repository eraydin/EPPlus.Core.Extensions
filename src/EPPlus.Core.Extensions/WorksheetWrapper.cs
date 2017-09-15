using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.RegularExpressions;

namespace EPPlus.Core.Extensions
{
    public class WorksheetWrapper<T>
    {
        internal string Name { get; set; }

        internal bool AppendHeaderRow { get; set; } = true;
        internal ExcelPackage Package { get; set; }
        internal IList<T> Rows { get; set; }
        internal IList<WorksheetColumn<T>> Columns { get; set; }
        internal IList<WorksheetTitleRow> Titles { get; set; }
        internal Action<ExcelColumn> ConfigureColumn { get; set; }
        internal Action<ExcelRange> ConfigureHeader { get; set; }
        internal Action<ExcelRange> ConfigureHeaderRow { get; set; }
        internal Action<ExcelRange, T> ConfigureCell { get; set; }

        /// <summary>
        /// Generates columns for all public properties on the type
        /// </summary>
        /// <returns></returns>
        internal IList<WorksheetColumn<T>> AutoGenerateColumns()
        {
            var columns = new List<WorksheetColumn<T>>();

            Type type = typeof(T);
            PropertyInfo[] properties = type.GetProperties();

            foreach (PropertyInfo property in properties)
            {
                var column = new WorksheetColumn<T>
                {
                    // Convert to sentence case
                    Header = Regex.Replace(property.Name, "[a-z][A-Z]", m => $"{m.Value[0]} {m.Value[1]}"),
                    Map = GetGetter<T>(property.Name),
                    ConfigureColumn = c => c.AutoFit()
                };
                columns.Add(column);
            }

            return columns;
        }

        /// <summary>
        /// Wraps creation of an Excel worksheet
        /// </summary>
        internal void AppendWorksheet()
        {
            if (Package == null)
            {
                Package = new ExcelPackage();
            }

            ExcelWorksheet worksheet = Package.Workbook.Worksheets.Add(this.Name);

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
                    worksheet.Cells[rowOffset + 1, i + 1].Style.Font.Bold = true;

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
                        worksheet.Cells[r + rowOffset + 1, c + 1].Value = Columns[c].Map(Rows[r]);

                        if (this.ConfigureCell != null)
                        {
                            this.ConfigureCell(worksheet.Cells[r + rowOffset + 1, c + 1], Rows[r]);
                        }
                        if (Columns[c].ConfigureCell != null)
                        {
                            Columns[c].ConfigureCell(worksheet.Cells[r + rowOffset + 1, c + 1], Rows[r]);
                        }
                    }
                }
            }

            //configure columns
            for (var i = 0; i < Columns.Count; i++)
            {
                if (ConfigureColumn != null)
                {
                    ConfigureColumn(worksheet.Column(i + 1));
                }
                if (Columns[i].ConfigureColumn != null)
                {
                    Columns[i].ConfigureColumn(worksheet.Column(i + 1));
                }
            }
        }

        /// <summary>
        /// Generates a Func from a propertyName
        /// </summary>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private Func<T, object> GetGetter<T>(string propertyName)
        {
            ParameterExpression arg = Expression.Parameter(typeof(T), "x");
            MemberExpression expression = Expression.Property(arg, propertyName);
            UnaryExpression conversion = Expression.Convert(expression, typeof(object));
            return Expression.Lambda<Func<T, object>>(conversion, arg).Compile();
        }
    }
}