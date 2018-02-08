using System;
using System.Linq;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace EPPlus.Core.Extensions
{
    public static class ExcelWorkbookExtensions
    {
        /// <summary>
        ///     Creates a named style on given Excel workbook
        /// </summary>
        /// <param name="workbook">The workbook</param>
        /// <param name="styleName">The name of style</param>
        /// <param name="style">The style actions will be applied</param>
        /// <returns></returns>
        public static ExcelWorkbook CreateNamedStyle(this ExcelWorkbook workbook, string styleName, Action<ExcelStyle> style)
        {
            if (workbook.Styles.NamedStyles.All(x => x.Name == styleName))
            {
                throw new ArgumentException($"The Excel package already has a style with the name of '{styleName}'");
            }

            ExcelNamedStyleXml errorStyle = workbook.Styles.CreateNamedStyle(styleName);
            style.Invoke(errorStyle.Style);

            return workbook;
        }
    }
}
