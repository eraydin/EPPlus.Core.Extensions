using System.Drawing;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.Core.Extensions.Style
{
    public static class ExcelWorksheetExtensions
    {
        /// <summary>
        ///     Sets the font of ExcelWorksheet cells from a Font object
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        public static ExcelWorksheet SetFont(this ExcelWorksheet worksheet, Font font)
        {
            worksheet.Cells.SetFont(font);
            return worksheet;
        }

        /// <summary>
        ///     Sets the font of given cell range from a Font object
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellRange"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        public static ExcelWorksheet SetFont(this ExcelWorksheet worksheet, ExcelRange cellRange, Font font)
        {
            worksheet.Cells[cellRange.Address].SetFont(font);
            return worksheet;
        }

        /// <summary>
        ///     Sets the font color of ExcelWorksheet cells from a Color object
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="fontColor"></param>
        /// <returns></returns>
        public static ExcelWorksheet SetFontColor(this ExcelWorksheet worksheet, Color fontColor)
        {
            worksheet.Cells.SetFontColor(fontColor);
            return worksheet;
        }

        /// <summary>
        ///     Sets the font color of given cell range from a Color object
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellRange"></param>
        /// <param name="fontColor"></param>
        /// <returns></returns>
        public static ExcelWorksheet SetFontColor(this ExcelWorksheet worksheet, ExcelRange cellRange, Color fontColor)
        {
            worksheet.Cells[cellRange.Address].SetFontColor(fontColor);
            return worksheet;
        }

        /// <summary>
        ///     Sets the background color of ExcelWorksheet cells from a Color object
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="fillStyle"></param>
        /// <returns></returns>
        public static ExcelWorksheet SetBackgroundColor(this ExcelWorksheet worksheet, Color backgroundColor, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            worksheet.Cells.SetBackgroundColor(backgroundColor, fillStyle);
            return worksheet;
        }

        /// <summary>
        ///     Sets the background color of given cell range from a Color object
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellRange"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="fillStyle"></param>
        /// <returns></returns>
        public static ExcelWorksheet SetBackgroundColor(this ExcelWorksheet worksheet, ExcelRange cellRange, Color backgroundColor, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            worksheet.Cells[cellRange.Address].SetBackgroundColor(backgroundColor, fillStyle);
            return worksheet;
        }

        /// <summary>
        ///     Sets the horizontal alignment of ExcelWorksheet cells
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="horizontalAlignment"></param>
        /// <returns></returns>
        public static ExcelWorksheet SetHorizontalAlignment(this ExcelWorksheet worksheet, ExcelHorizontalAlignment horizontalAlignment) => worksheet.SetHorizontalAlignment(worksheet.Cells, horizontalAlignment);

        /// <summary>
        ///     Sets the horizontal alignment of given cell range
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellRange"></param>
        /// <param name="horizontalAlignment"></param>
        /// <returns></returns>
        public static ExcelWorksheet SetHorizontalAlignment(this ExcelWorksheet worksheet, ExcelRange cellRange, ExcelHorizontalAlignment horizontalAlignment)
        {
            worksheet.Cells[cellRange.Address].SetHorizontalAlignment(horizontalAlignment);
            return worksheet;
        }

        /// <summary>
        ///     Sets the vertical alignment of ExcelWorksheet cells
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="verticalAlignment"></param>
        /// <returns></returns>
        public static ExcelWorksheet SetVerticalAlignment(this ExcelWorksheet worksheet, ExcelVerticalAlignment verticalAlignment) => worksheet.SetVerticalAlignment(worksheet.Cells, verticalAlignment);

        /// <summary>
        ///     Sets the vertical alignment of given cell range
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellRange"></param>
        /// <param name="verticalAlignment"></param>
        /// <returns></returns>
        public static ExcelWorksheet SetVerticalAlignment(this ExcelWorksheet worksheet, ExcelRange cellRange, ExcelVerticalAlignment verticalAlignment)
        {
            worksheet.Cells[cellRange.Address].SetVerticalAlignment(verticalAlignment);
            return worksheet;
        }
    }
}
