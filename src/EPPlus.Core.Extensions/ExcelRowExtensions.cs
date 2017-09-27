using System.Drawing;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.Core.Extensions
{
    public static class ExcelRowExtensions
    {
        public static ExcelRow SetFont(this ExcelRow row, Font font)
        {
            row.Style.Font.SetFromFont(font);
            return row;
        }

        public static ExcelRow SetFontColor(this ExcelRow row, Color fontColor)
        {
            row.Style.Font.Color.SetColor(fontColor);
            return row;
        }

        public static ExcelRow SetBackgroundColor(this ExcelRow row, Color backgroundColor)
        {
            row.Style.Fill.BackgroundColor.SetColor(backgroundColor);
            return row;
        }

        public static ExcelRow SetHorizontalAlignment(this ExcelRow row, ExcelHorizontalAlignment horizontalAlignment)
        {
            row.Style.HorizontalAlignment = horizontalAlignment;
            return row;
        }

        public static ExcelRow SetVerticalAlignment(this ExcelRow row, ExcelVerticalAlignment verticalAlignment)
        {
            row.Style.VerticalAlignment = verticalAlignment;
            return row;
        }
    }
}
