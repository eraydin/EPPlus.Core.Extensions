using System.Drawing;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.Core.Extensions
{
    public static class ExcelRangeExtensions
    {
        /// <summary>
        ///     Sets the font of ExcelRangeBase from a Font object
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
        ///     Sets the font color of ExcelRangeBase from a Color object
        /// </summary>
        /// <param name="range"></param>
        /// <param name="fontColor"></param>
        /// <returns></returns>
        public static ExcelRangeBase SetFontColor(this ExcelRangeBase range, Color fontColor)
        {
            range.Style.Font.Color.SetColor(fontColor);
            return range;
        }

        /// <summary>
        ///     Sets the background color of ExcelRangeBase from a Color object
        /// </summary>
        /// <param name="range"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="fillStyle"></param>
        /// <returns></returns>
        public static ExcelRangeBase SetBackgroundColor(this ExcelRangeBase range, Color backgroundColor, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            range.Style.Fill.PatternType = fillStyle;
            range.Style.Fill.BackgroundColor.SetColor(backgroundColor);
            return range;
        }

        /// <summary>
        ///     Sets the horizontal alignment of ExcelRangeBase
        /// </summary>
        /// <param name="range"></param>
        /// <param name="horizontalAlignment"></param>
        /// <returns></returns>
        public static ExcelRangeBase SetHorizontalAlignment(this ExcelRangeBase range, ExcelHorizontalAlignment horizontalAlignment)
        {
            range.Style.HorizontalAlignment = horizontalAlignment;
            return range;
        }

        /// <summary>
        ///     Sets the vertical alignment of ExcelRangeBase
        /// </summary>
        /// <param name="range"></param>
        /// <param name="verticalAlignment"></param>
        /// <returns></returns>
        public static ExcelRangeBase SetVerticalAlignment(this ExcelRangeBase range, ExcelVerticalAlignment verticalAlignment)
        {
            range.Style.VerticalAlignment = verticalAlignment;
            return range;
        }
    }
}
