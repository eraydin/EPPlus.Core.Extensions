using System.Drawing;

using OfficeOpenXml.Style;

namespace EPPlus.Core.Extensions.Style
{
    public static class ExcelStyleExtensions
    {
        /// <summary>
        ///     Sets border style around the range.
        /// </summary>
        /// <param name="thisStyle"></param>
        /// <param name="style"></param>
        /// <returns></returns>
        public static ExcelStyle BorderAround(this ExcelStyle thisStyle, ExcelBorderStyle style)
        {
            thisStyle.BorderAround(style, Color.Black);
            return thisStyle;
        }

        /// <summary>
        ///     Sets border style around the range.
        /// </summary>
        /// <param name="thisStyle"></param>
        /// <param name="borderStyle"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static ExcelStyle BorderAround(this ExcelStyle thisStyle, ExcelBorderStyle borderStyle, Color color)
        {
            thisStyle.Border.Top.Style = borderStyle;
            thisStyle.Border.Left.Style = borderStyle;
            thisStyle.Border.Bottom.Style = borderStyle;
            thisStyle.Border.Right.Style = borderStyle;

            thisStyle.Border.Top.Color.SetColor(color);
            thisStyle.Border.Left.Color.SetColor(color);
            thisStyle.Border.Bottom.Color.SetColor(color);
            thisStyle.Border.Right.Color.SetColor(color);
            return thisStyle;
        }

        /// <summary>
        ///     Sets background color of Excel style
        /// </summary>
        /// <param name="thisStyle">The Excel style</param>
        /// <param name="color">The color</param>
        /// <param name="fillStyle">The fill style of background</param>
        /// <returns></returns>
        public static ExcelStyle SetBackgroundColor(this ExcelStyle thisStyle, Color color, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            thisStyle.Fill.PatternType = fillStyle;
            thisStyle.Fill.BackgroundColor.SetColor(color);
            return thisStyle;
        }

        /// <summary>
        ///     Sets font of Excel style
        /// </summary>
        /// <param name="thisStyle">The Excel style</param>
        /// <param name="font">The font</param>
        /// <returns></returns>
        public static ExcelStyle SetFont(this ExcelStyle thisStyle, Font font)
        {
            thisStyle.Font.SetFromFont(font);
            return thisStyle;
        }

        /// <summary>
        ///     Sets font and color of Excel style
        /// </summary>
        /// <param name="thisStyle">The Excel style</param>
        /// <param name="font">The font</param>
        /// <param name="color">The color</param>
        /// <returns></returns>
        public static ExcelStyle SetFont(this ExcelStyle thisStyle, Font font, Color color)
        {
            thisStyle.Font.SetFromFont(font);
            thisStyle.SetFontColor(color);
            return thisStyle;
        }

        /// <summary>
        ///     Sets font color of Excel style
        /// </summary>
        /// <param name="thisStyle"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static ExcelStyle SetFontColor(this ExcelStyle thisStyle, Color color)
        {
            thisStyle.Font.Color.SetColor(color);
            return thisStyle;
        }

        public static ExcelStyle SetFontName(this ExcelStyle thisStyle, string newFontName)
        {
            thisStyle.Font.Name = newFontName;
            return thisStyle;
        }

        public static ExcelStyle SetFontAsBold(this ExcelStyle thisStyle)
        {
            thisStyle.Font.Bold = true;
            return thisStyle;
        }

        /// <summary>
        ///     Sets horizontal alignment of Excel style
        /// </summary>
        /// <param name="thisStyle">The Excel style</param>
        /// <param name="horizontalAlignment"></param>
        /// <returns></returns>
        public static ExcelStyle SetHorizontalAlignment(this ExcelStyle thisStyle, ExcelHorizontalAlignment horizontalAlignment)
        {
            thisStyle.HorizontalAlignment = horizontalAlignment;
            return thisStyle;
        }

        /// <summary>
        ///     Sets vertical alignment of Excel style
        /// </summary>
        /// <param name="thisStyle">The Excel style</param>
        /// <param name="verticalAlignment"></param>
        /// <returns></returns>
        public static ExcelStyle SetVerticalAlignment(this ExcelStyle thisStyle, ExcelVerticalAlignment verticalAlignment)
        {
            thisStyle.VerticalAlignment = verticalAlignment;
            return thisStyle;
        }
    }
}
