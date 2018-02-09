using System.Drawing;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.Core.Extensions
{
    public static class ExcelRangeExtensions
    {
        /// <summary>
        ///     Sets the border style of given range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="style"></param>
        /// <returns></returns>
        public static ExcelRangeBase BorderAround(this ExcelRangeBase range, ExcelBorderStyle style)
        {
            range.Style.BorderAround(style);
            return range;
        }

        /// <summary>
        ///     Sets the border style and color of given range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static ExcelRangeBase BorderAround(this ExcelRangeBase range, ExcelBorderStyle style, Color color)
        {
            range.Style.BorderAround(style, color);
            return range;
        }

        /// <summary>
        ///     Sets the background color of given range from a Color object
        /// </summary>
        /// <param name="range"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="fillStyle"></param>
        /// <returns></returns>
        public static ExcelRangeBase SetBackgroundColor(this ExcelRangeBase range, Color backgroundColor, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            range.Style.SetBackgroundColor(backgroundColor, fillStyle);
            return range;
        }

        /// <summary>
        ///     Sets the border color of given range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static ExcelRangeBase SetBorderColor(this ExcelRangeBase range, Color color)
        {
            range.BorderAround(ExcelBorderStyle.Thin, color);
            return range;
        }

        /// <summary>
        ///     Sets the font of given range from a Font object
        /// </summary>
        /// <param name="range"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        public static ExcelRangeBase SetFont(this ExcelRangeBase range, Font font)
        {
            range.Style.Font.SetFromFont(font);
            return range;
        }

        /// <summary>
        ///     Sets the font color of given range from a Color object
        /// </summary>
        /// <param name="range"></param>
        /// <param name="fontColor"></param>
        /// <returns></returns>
        public static ExcelRangeBase SetFontColor(this ExcelRangeBase range, Color fontColor)
        {
            range.Style.SetFontColor(fontColor);
            return range;
        }

        /// <summary>
        ///     Sets the horizontal alignment of given range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="horizontalAlignment"></param>
        /// <returns></returns>
        public static ExcelRangeBase SetHorizontalAlignment(this ExcelRangeBase range, ExcelHorizontalAlignment horizontalAlignment)
        {
            range.Style.SetHorizontalAlignment(horizontalAlignment);
            return range;
        }

        /// <summary>
        ///     Sets the vertical alignment of given range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="verticalAlignment"></param>
        /// <returns></returns>
        public static ExcelRangeBase SetVerticalAlignment(this ExcelRangeBase range, ExcelVerticalAlignment verticalAlignment)
        {
            range.Style.SetVerticalAlignment(verticalAlignment);
            return range;
        }
    }
}
