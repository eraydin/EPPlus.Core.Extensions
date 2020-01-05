using System.Drawing;
using System.Linq;

using EPPlus.Core.Extensions.Style;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using Xunit;

namespace EPPlus.Core.Extensions.Tests.StyleTests
{
    public class ExcelWorksheetTests : TestBase
    {
        [Fact]
        public void Should_change_background_color_of_specific_range_of_worksheet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet(2);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.SetBackgroundColor(Color.Yellow);
            worksheet.SetBackgroundColor(worksheet.Cells[1, 3, 1, 3], Color.Brown, ExcelFillStyle.DarkTrellis);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Cells[1, 3, 1, 3].Style.Fill.PatternType.Should().Be(ExcelFillStyle.DarkTrellis);
            worksheet.Cells[1, 3, 1, 3].Style.Fill.BackgroundColor.Rgb.Should().Be($"{Color.Brown.ToArgb() & 0xFFFFFFFF:X8}");
            worksheet.Cells[2, 3, 2, 3].Style.Fill.BackgroundColor.Rgb.Should().Be($"{Color.Yellow.ToArgb() & 0xFFFFFFFF:X8}");
        }

        [Fact]
        public void Should_change_background_color_of_the_worksheet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet(2);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.SetBackgroundColor(Color.Brown, ExcelFillStyle.DarkTrellis);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Cells.Style.Fill.PatternType.Should().Be(ExcelFillStyle.DarkTrellis);
            worksheet.Cells.Style.Fill.BackgroundColor.Rgb.Should().Be($"{Color.Brown.ToArgb() & 0xFFFFFFFF:X8}");
        }

        [Fact]
        public void Should_change_font_color_of_specific_range_of_worksheet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet(3);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.SetFontColor(Color.Yellow);
            worksheet.SetFontColor(worksheet.Cells[1, 2, 1, 3], Color.BlueViolet);
            worksheet.SetFont(new Font("Verdana", 12));

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Cells[1, 2, 1, 3].Style.Font.Color.Rgb.Should().Be($"{Color.BlueViolet.ToArgb() & 0xFFFFFFFF:X8}");
            worksheet.Cells[2, 2, 2, 3].Style.Font.Color.Rgb.Should().Be($"{Color.Yellow.ToArgb() & 0xFFFFFFFF:X8}");
            worksheet.Cells.Style.Font.Name.Should().Be("Verdana");
        }

        [Fact]
        public void Should_change_font_color_of_the_worksheet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet(3);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.SetFontColor(Color.BlueViolet);
            worksheet.SetFont(new Font("Verdana", 12));

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Cells.Style.Font.Color.Rgb.Should().Be($"{Color.BlueViolet.ToArgb() & 0xFFFFFFFF:X8}");
            worksheet.Cells.Style.Font.Name.Should().Be("Verdana");
        }

        [Fact]
        public void Should_change_font_of_specific_range_of_worksheet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet(3);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.SetFontColor(Color.BlueViolet);
            worksheet.SetFont(new Font("Arial", 12));
            worksheet.SetFont(worksheet.Cells[1, 2, 1, 2], new Font("Verdana", 12));

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Cells.Style.Font.Color.Rgb.Should().Be($"{Color.BlueViolet.ToArgb() & 0xFFFFFFFF:X8}");
            worksheet.Cells[1, 2, 1, 2].Style.Font.Name.Should().Be("Verdana");
            worksheet.Cells[2, 2, 2, 2].Style.Font.Name.Should().Be("Arial");
        }

        [Fact]
        public void Should_change_font_of_the_worksheet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.Workbook.Worksheets.First();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.SetFont(new Font("Arial", 15));

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Cells.Style.Font.Name.Should().Be("Arial");
            worksheet.Cells.Style.Font.Size.Should().Be(15);
        }


        [Fact]
        public void Should_set_horizontal_alignment_of_the_worksheet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.GetWorksheet(4);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.SetHorizontalAlignment(ExcelHorizontalAlignment.Distributed);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Cells.Style.HorizontalAlignment.Should().Be(ExcelHorizontalAlignment.Distributed);
        }

        [Fact]
        public void Should_set_vertical_alignment_of_the_worksheet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = ExcelPackage1.Workbook.Worksheets.Last();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.SetVerticalAlignment(ExcelVerticalAlignment.Justify);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Cells.Style.VerticalAlignment.Should().Be(ExcelVerticalAlignment.Justify);
        }
    }
}