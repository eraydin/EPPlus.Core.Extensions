using System.Drawing;

using OfficeOpenXml.Style;

namespace EPPlus.Core.Extensions
{
    public static class ExcelStyleExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="thisStyle"></param>
        /// <param name="borderStyle"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static ExcelStyle BorderAround(this ExcelStyle thisStyle, ExcelBorderStyle borderStyle, Color color)
        {
            thisStyle.Border.Bottom.Style = borderStyle;
            thisStyle.Border.Top.Style = borderStyle;
            thisStyle.Border.Right.Style = borderStyle;
            thisStyle.Border.Left.Style = borderStyle;

            thisStyle.Border.Bottom.Color.SetColor(color);
            thisStyle.Border.Top.Color.SetColor(color);
            thisStyle.Border.Right.Color.SetColor(color);
            thisStyle.Border.Left.Color.SetColor(color);
            return thisStyle;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="thisStyle"></param>
        /// <param name="color"></param>
        /// <param name="fillStyle"></param>
        /// <returns></returns>
        public static ExcelStyle SetBackgroundColor(this ExcelStyle thisStyle, Color color, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            thisStyle.Fill.PatternType = fillStyle;
            thisStyle.Fill.BackgroundColor.SetColor(color);
            return thisStyle;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="thisStyle"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        public static ExcelStyle SetFont(this ExcelStyle thisStyle, Font font)
        {
            thisStyle.Font.SetFromFont(font);
            return thisStyle;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="thisStyle"></param>
        /// <param name="font"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static ExcelStyle SetFont(this ExcelStyle thisStyle, Font font, Color color)
        {
            thisStyle.Font.SetFromFont(font);
            thisStyle.SetFontColor(color);
            return thisStyle;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="thisStyle"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static ExcelStyle SetFontColor(this ExcelStyle thisStyle, Color color)
        {
            thisStyle.Font.Color.SetColor(color);
            return thisStyle;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="thisStyle"></param>
        /// <param name="horizontalAlignment"></param>
        /// <returns></returns>
        public static ExcelStyle SetHorizontalAlignment(this ExcelStyle thisStyle, ExcelHorizontalAlignment horizontalAlignment)
        {
            thisStyle.HorizontalAlignment = horizontalAlignment;
            return thisStyle;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="thisStyle"></param>
        /// <param name="verticalAlignment"></param>
        /// <returns></returns>
        public static ExcelStyle SetVerticalAlignment(this ExcelStyle thisStyle, ExcelVerticalAlignment verticalAlignment)
        {
            thisStyle.VerticalAlignment = verticalAlignment;
            return thisStyle;
        }
    }
}
