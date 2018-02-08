using System.Drawing;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelWorkbook_Tests : TestBase
    {
        [Fact]
        public void Should_create_a_named_style_with_style_actions()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorkbook workbook = excelPackage.Workbook;
            ExcelWorksheet firstWorksheet = workbook.Worksheets[1];
            ExcelWorksheet secondWorksheet = workbook.Worksheets[2];
            var nameOfStyle = "NamedStyle1";
            var fontName = "Arial";

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            workbook.CreateNamedStyle(nameOfStyle, style =>
            {
                style.SetBackgroundColor(Color.Blue, ExcelFillStyle.DarkDown);
                style.SetFont(new Font(fontName, 12, FontStyle.Bold), Color.Yellow);
            });

            firstWorksheet.Cells[1, 1, 1, 1].StyleName = nameOfStyle;
            secondWorksheet.Cells[secondWorksheet.GetValuedDimension().Address].StyleName = nameOfStyle;

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            firstWorksheet.Cells[1, 1, 1, 1].StyleName.Should().Be(nameOfStyle);
            firstWorksheet.Cells[1, 1, 1, 1].Style.Font.Name.Should().Be(fontName);
            firstWorksheet.Cells[1, 1, 1, 1].Style.Font.Color.Rgb.Should().Be(string.Format("{0:X8}", Color.Yellow.ToArgb() & 0xFFFFFFFF));
            firstWorksheet.Cells[1, 1, 1, 1].Style.Fill.BackgroundColor.Rgb.Should().Be(string.Format("{0:X8}", Color.Blue.ToArgb() & 0xFFFFFFFF));
            firstWorksheet.Cells[1, 1, 1, 1].Style.Fill.PatternType.Should().Be(ExcelFillStyle.DarkDown);

            secondWorksheet.Cells[secondWorksheet.GetValuedDimension().Address].StyleName.Should().Be(nameOfStyle);
            secondWorksheet.Cells[secondWorksheet.GetValuedDimension().Address].Style.Font.Name.Should().Be(fontName);
            secondWorksheet.Cells[secondWorksheet.GetValuedDimension().Address].Style.Font.Color.Rgb.Should().Be(string.Format("{0:X8}", Color.Yellow.ToArgb() & 0xFFFFFFFF));
            secondWorksheet.Cells[secondWorksheet.GetValuedDimension().Address].Style.Fill.BackgroundColor.Rgb.Should().Be(string.Format("{0:X8}", Color.Blue.ToArgb() & 0xFFFFFFFF));
            secondWorksheet.Cells[secondWorksheet.GetValuedDimension().Address].Style.Fill.PatternType.Should().Be(ExcelFillStyle.DarkDown);
        }
    }
}
