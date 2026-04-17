using System.Drawing;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.Core.Extensions.Style
{
    public static class ExcelColumnExtensions
    {
        /// <summary>
        ///     Sets the font of ExcelColumn from font name and size
        /// </summary>
        /// <param name="column"></param>
        /// <param name="fontName"></param>
        /// <param name="size"></param>
        /// <param name="bold">Bold</param>
        /// <param name="italic">Italic</param>
        /// <returns></returns>
        public static ExcelColumn SetFont(this ExcelColumn column, string fontName, float size, bool bold = false, bool italic = false)
        {
            column.Style.SetFont(fontName, size, bold, italic);
            return column;
        }

        /// <summary>
        ///     Sets the font color of ExcelColumn
        /// </summary>
        /// <param name="column"></param>
        /// <param name="fontColor"></param>
        /// <returns></returns>
        public static ExcelColumn SetFontColor(this ExcelColumn column, Color fontColor)
        {
            column.Style.SetFontColor(fontColor);
            return column;
        }
    
        /// <summary>
        ///     Sets the background color of ExcelColumn from a Color object
        /// </summary>
        /// <param name="column"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="fillStyle"></param>
        /// <returns></returns>
        public static ExcelColumn SetBackgroundColor(this ExcelColumn column, Color backgroundColor, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            column.Style.SetBackgroundColor(backgroundColor, fillStyle);
            return column;
        }

        /// <summary>
        ///     Sets the horizontal alignment of ExcelColumn
        /// </summary>
        /// <param name="column"></param>
        /// <param name="horizontalAlignment"></param>
        /// <returns></returns>
        public static ExcelColumn SetHorizontalAlignment(this ExcelColumn column, ExcelHorizontalAlignment horizontalAlignment)
        {
            column.Style.SetHorizontalAlignment(horizontalAlignment);
            return column;
        }

        /// <summary>
        ///     Sets the vertical alignment of ExcelColumn
        /// </summary>
        /// <param name="column"></param>
        /// <param name="verticalAlignment"></param>
        /// <returns></returns>
        public static ExcelColumn SetVerticalAlignment(this ExcelColumn column, ExcelVerticalAlignment verticalAlignment)
        {
            column.Style.SetVerticalAlignment(verticalAlignment);
            return column;
        }
    }
}
