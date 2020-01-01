using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

using EPPlus.Core.Extensions.Configuration;
using EPPlus.Core.Extensions.Enrichments;
using EPPlus.Core.Extensions.Helpers;

using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.Core.Extensions
{
    public class WorksheetWrapper<T>
    {
        internal string Name { get; set; }

        internal bool AppendHeaderRow { get; set; } = true;

        internal ExcelPackage Package { get; set; }

        internal IEnumerable<T> Rows { get; set; }

        internal List<WorksheetColumn<T>> Columns { get; set; }

        internal List<WorksheetTitleRow> Titles { get; set; }

        internal ExcelCreateConfiguration<T> Configuration { get; } = DefaultExcelCreateConfiguration<T>.Instance;

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
                    ExcelRange range = worksheet.Cells[rowOffset + 1 + i, 1, rowOffset + 1 + i, Columns.Count];
                    range.Merge = true;
                    range.Value = Titles[i].Title;

                    Configuration.ConfigureTitle?.Invoke(range);
                    Titles[i].ConfigureTitle?.Invoke(range);
                }

                rowOffset = rowOffset + Titles.Count;
            }

            //render headers
            if (AppendHeaderRow)
            {
                for (var i = 0; i < Columns.Count; i++)
                {
                    worksheet.Cells[rowOffset + 1, i + 1].Value = Columns[i].Header;
                    Configuration.ConfigureHeader?.Invoke(worksheet.Cells[rowOffset + 1, i + 1]);
                }

                //configure the header row
                if (Configuration.ConfigureHeaderRow != null)
                {
                    Configuration.ConfigureHeaderRow.Invoke(worksheet.Cells[rowOffset + 1, 1, rowOffset + 1, Columns.Count]);
                }

                rowOffset++;
            }

            CreateTableIfPossible(worksheet);

            //render data
            if (Rows != null)
            {
                for (var r = 0; r < Rows.Count(); r++)
                {
                    for (var c = 0; c < Columns.Count(); c++)
                    {
                        worksheet.Cells[r + rowOffset + 1, c + 1].Value = Columns[c].Map(Rows.ElementAt(r));

                        Configuration.ConfigureCell?.Invoke(worksheet.Cells[r + rowOffset + 1, c + 1], Rows.ElementAt(r));
                    }
                }
            }

            //configure columns
            for (var i = 0; i < Columns.Count; i++)
            {
                Configuration.ConfigureColumn?.Invoke(worksheet.Column(i + 1));
                Columns[i].ConfigureColumn?.Invoke(worksheet.Column(i + 1));
            }
        }
     
        /// <summary>
        ///     Generates columns for all public properties on the type
        /// </summary>
        /// <returns></returns>
        private List<WorksheetColumn<T>> AutoGenerateColumns()
        {
            List<ColumnAttributeAndPropertyInfo> propertyInfoAndColumnAttributes = typeof(T).GetExcelTableColumnAttributesWithPropertyInfo();

            var columns = propertyInfoAndColumnAttributes.Select(x => new WorksheetColumn<T>
                                                         {
                                                             Header = x.ToString(),
                                                             Map = GetGetter<T>(x.PropertyInfo.Name)
                                                         })
                                                         .ToList();
            return columns;
        }

        private void CreateTableIfPossible(ExcelWorksheet worksheet)
        {
            try
            {
                var tableStartRow = (Titles?.Count ?? 0) + 1;
                var tableEndRow = tableStartRow + (Rows?.Count() ?? 0);
                var columnsCount = (Columns?.Count ?? 1);
                var tableRange = worksheet.Cells[tableStartRow, 1, tableEndRow, columnsCount];
                worksheet.Tables.Add(tableRange, StringHelper.GenerateRandomTableName()).TableStyle = TableStyles.None;
            }
            catch
            {
                // ignored
            }
        }

        private Func<TP, object> GetGetter<TP>(string propertyName)
        {
            ParameterExpression arg = Expression.Parameter(typeof(TP), "x");
            MemberExpression expression = Expression.Property(arg, propertyName);
            UnaryExpression conversion = Expression.Convert(expression, typeof(object));
            return Expression.Lambda<Func<TP, object>>(conversion, arg).Compile();
        }
    }
}
