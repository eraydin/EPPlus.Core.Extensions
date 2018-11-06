using System;
using System.Linq;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

using static EPPlus.Core.Extensions.Helpers.Guard;

namespace EPPlus.Core.Extensions.Style
{
    public static class ExcelWorkbookExtensions
    {
        /// <summary>
        ///     Creates a named style on the Excel workbook. If the named style is already exists then throws the <see cref="ArgumentException"/>
        /// </summary>
        /// <param name="workbook">The workbook</param>
        /// <param name="styleName">The name of style</param>
        /// <param name="styleAction">The style actions which will be applied</param>
        /// <returns></returns>
        public static ExcelWorkbook CreateNamedStyle(this ExcelWorkbook workbook, string styleName, Action<ExcelStyle> styleAction)
        {
            NotNull(styleAction, nameof(styleAction));
            ThrowIfConditionMet(workbook.Styles.NamedStyles.Any(x => x.Name == styleName), "The Excel package already has a style with the name of '{0}'", styleName);
            
            ExcelNamedStyleXml errorStyle = workbook.Styles.CreateNamedStyle(styleName);
            styleAction.Invoke(errorStyle.Style);

            return workbook;
        }

        /// <summary>
        ///     Creates a named style if the given name is not exists on the Excel workbook
        /// </summary>
        /// <param name="workbook">The workbook</param>
        /// <param name="styleName">The name of style</param>
        /// <param name="styleAction">The style action which will be applied</param>
        /// <returns></returns>
        public static ExcelWorkbook CreateNamedStyleIfNotExists(this ExcelWorkbook workbook, string styleName, Action<ExcelStyle> styleAction)
        {
            NotNull(styleAction, nameof(styleAction));

            if (workbook.Styles.NamedStyles.All(x => x.Name != styleName))
            {
                ExcelNamedStyleXml errorStyle = workbook.Styles.CreateNamedStyle(styleName);
                styleAction.Invoke(errorStyle.Style);
            }
            return workbook;
        }
    }
}
