using System.Drawing;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.Core.Extensions.Style
{
    public static class ExcelRowExtensions
    {
        /// <summary>
        ///     Sets the font of ExcelRow from a Font object
        /// </summary>
        /// <param name="row"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        public static ExcelRow SetFont(this ExcelRow row, Font font)
        {
            row.Style.SetFont(font);
            return row;
        }

        /// <summary>
        ///     Sets the font color of ExcelRow from a Color object
        /// </summary>
        /// <param name="row"></param>
        /// <param name="fontColor"></param>
        /// <returns></returns>
        public static ExcelRow SetFontColor(this ExcelRow row, Color fontColor)
        {
            row.Style.SetFontColor(fontColor);
            return row;
        }

        /// <summary>
        ///     Sets the background color of ExcelRow from a Color object
        /// </summary>
        /// <param name="row"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="fillStyle"></param>
        /// <returns></returns>
        public static ExcelRow SetBackgroundColor(this ExcelRow row, Color backgroundColor, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            row.Style.SetBackgroundColor(backgroundColor, fillStyle);
            return row;
        }

        /// <summary>
        ///     Sets the horizontal alignment of ExcelRow
        /// </summary>
        /// <param name="row"></param>
        /// <param name="horizontalAlignment"></param>
        /// <returns></returns>
        public static ExcelRow SetHorizontalAlignment(this ExcelRow row, ExcelHorizontalAlignment horizontalAlignment)
        {
            row.Style.SetHorizontalAlignment(horizontalAlignment);
            return row;
        }

        /// <summary>
        ///     Sets the vertical alignment of ExcelRow
        /// </summary>
        /// <param name="row"></param>
        /// <param name="verticalAlignment"></param>
        /// <returns></returns>
        public static ExcelRow SetVerticalAlignment(this ExcelRow row, ExcelVerticalAlignment verticalAlignment)
        {
            row.Style.SetVerticalAlignment(verticalAlignment);
            return row;
        }
    }
}
