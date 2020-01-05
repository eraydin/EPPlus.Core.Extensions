using System.Drawing;
using System.Linq;

using EPPlus.Core.Extensions.Style;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using Xunit;

namespace EPPlus.Core.Extensions.Tests.StyleTests
{
    public class ExcelRangeExtensionsTests : TestBase
    {
        [Fact]
        public void Should_change_background_color_of_the_range()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelRange forthColumn = ExcelPackage1.Workbook.Worksheets.First().Cells[1, 4];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.SetBackgroundColor(Color.Brown, ExcelFillStyle.DarkTrellis);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.Style.Fill.PatternType.Should().Be(ExcelFillStyle.DarkTrellis);
            forthColumn.Style.Fill.BackgroundColor.Rgb.Should().Be($"{Color.Brown.ToArgb() & 0xFFFFFFFF:X8}");
        }

        [Fact]
        public void Should_change_border_color_of_given_cell_range()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelRange forthColumn = ExcelPackage1.Workbook.Worksheets.First().Cells[1, 4];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.SetBorderColor(Color.Purple);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.Style.Border.Left.Color.Rgb.Should().Be($"{Color.Purple.ToArgb() & 0xFFFFFFFF:X8}");
            forthColumn.Style.Border.Left.Color.Rgb.Should().Be($"{Color.Purple.ToArgb() & 0xFFFFFFFF:X8}");
            forthColumn.Style.Border.Right.Color.Rgb.Should().Be($"{Color.Purple.ToArgb() & 0xFFFFFFFF:X8}");
            forthColumn.Style.Border.Top.Color.Rgb.Should().Be($"{Color.Purple.ToArgb() & 0xFFFFFFFF:X8}");
            forthColumn.Style.Border.Bottom.Color.Rgb.Should().Be($"{Color.Purple.ToArgb() & 0xFFFFFFFF:X8}");
        }

        [Fact]
        public void Should_change_both_border_style_and_border_color_of_given_cell_range()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelRange forthColumn = ExcelPackage1.Workbook.Worksheets.First().Cells[1, 4];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.BorderAround(ExcelBorderStyle.Dashed, Color.Red);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.Style.Border.Left.Style.Should().Be(ExcelBorderStyle.Dashed);
            forthColumn.Style.Border.Right.Style.Should().Be(ExcelBorderStyle.Dashed);
            forthColumn.Style.Border.Top.Style.Should().Be(ExcelBorderStyle.Dashed);
            forthColumn.Style.Border.Bottom.Style.Should().Be(ExcelBorderStyle.Dashed);

            forthColumn.Style.Border.Left.Color.Rgb.Should().Be($"{Color.Red.ToArgb() & 0xFFFFFFFF:X8}");
            forthColumn.Style.Border.Left.Color.Rgb.Should().Be($"{Color.Red.ToArgb() & 0xFFFFFFFF:X8}");
            forthColumn.Style.Border.Right.Color.Rgb.Should().Be($"{Color.Red.ToArgb() & 0xFFFFFFFF:X8}");
            forthColumn.Style.Border.Top.Color.Rgb.Should().Be($"{Color.Red.ToArgb() & 0xFFFFFFFF:X8}");
            forthColumn.Style.Border.Bottom.Color.Rgb.Should().Be($"{Color.Red.ToArgb() & 0xFFFFFFFF:X8}");
        }

        [Fact]
        public void Should_change_font_color_of_the_range()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelRange forthColumn = ExcelPackage1.Workbook.Worksheets.First().Cells[1, 4];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.SetFontColor(Color.BlueViolet);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.Style.Font.Color.Rgb.Should().Be($"{Color.BlueViolet.ToArgb() & 0xFFFFFFFF:X8}");
        }

        [Fact]
        public void Should_change_font_of_the_range()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelRange forthColumn = ExcelPackage1.Workbook.Worksheets.First().Cells[1, 4];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.SetFont(new Font("Arial", 15));

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.Style.Font.Name.Should().Be("Arial");
            forthColumn.Style.Font.Size.Should().Be(15);
        }

        [Fact]
        public void Should_change_set_border_style_of_given_cell_range()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelRange forthColumn = ExcelPackage1.Workbook.Worksheets.First().Cells[1, 4];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.BorderAround(ExcelBorderStyle.Dotted);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.Style.Border.Left.Style.Should().Be(ExcelBorderStyle.Dotted);
            forthColumn.Style.Border.Right.Style.Should().Be(ExcelBorderStyle.Dotted);
            forthColumn.Style.Border.Top.Style.Should().Be(ExcelBorderStyle.Dotted);
            forthColumn.Style.Border.Bottom.Style.Should().Be(ExcelBorderStyle.Dotted);
        }

        [Fact]
        public void Should_set_horizontal_alignment_of_the_range()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelRange forthColumn = ExcelPackage1.Workbook.Worksheets.First().Cells[1, 4];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.SetHorizontalAlignment(ExcelHorizontalAlignment.Distributed);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.Style.HorizontalAlignment.Should().Be(ExcelHorizontalAlignment.Distributed);
        }

        [Fact]
        public void Should_set_vertical_alignment_of_the_range()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelRange forthColumn = ExcelPackage1.Workbook.Worksheets.First().Cells[1, 4];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.SetVerticalAlignment(ExcelVerticalAlignment.Justify);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.Style.VerticalAlignment.Should().Be(ExcelVerticalAlignment.Justify);
        }
    }
}
