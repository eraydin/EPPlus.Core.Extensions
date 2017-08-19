using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace EPPlus.Core.Extensions
{
    public static class ToExcelExtensions
    {
        public static ExcelPackage ToPackage<T>(this IList<T> rows)
        {
            throw new NotImplementedException();
        }

        public static byte[] ToXlsx<T>(this IList<T> rows)
        {
            throw new NotImplementedException();
        }

        public static ExcelWorksheet ToWorksheet<T>(this IList<T> rows, string name)
        {
            throw new NotImplementedException();
        }

        public static ExcelWorksheet WithColumn<T>(this ExcelWorksheet worksheet, Func<T, object> map)
        {
            throw new NotImplementedException();
        }

        public static ExcelWorksheet WithTitle<T>(this ExcelWorksheet worksheet, string title)
        {
            throw new NotImplementedException();
        }

        public static ExcelWorksheet AppendWorksheet<T>(this ExcelWorksheet previousSheet, IList<T> rows, string name)
        {
            throw new NotImplementedException();
        }
    }
}
