using System.Drawing;
using System.Linq;

using EPPlus.Core.Extensions.Style;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using Xunit;

namespace EPPlus.Core.Extensions.Tests.StyleTests
{
    public class ExcelColumnExtensions_Tests : TestBase
    {
        [Fact]
        public void Should_change_background_color_of_the_column()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelColumn forthColumn = excelPackage.Workbook.Worksheets.First().Column(4);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.SetBackgroundColor(Color.Brown, ExcelFillStyle.DarkTrellis);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.Style.Fill.PatternType.Should().Be(ExcelFillStyle.DarkTrellis);
            forthColumn.Style.Fill.BackgroundColor.Rgb.Should().Be(string.Format("{0:X8}", Color.Brown.ToArgb() & 0xFFFFFFFF));
        }

        [Fact]
        public void Should_change_font_color_of_the_column()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelColumn forthColumn = excelPackage.Workbook.Worksheets.First().Column(4);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.SetFontColor(Color.BlueViolet);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            forthColumn.Style.Font.Color.Rgb.Should().Be(string.Format("{0:X8}", Color.BlueViolet.ToArgb() & 0xFFFFFFFF));
        }

        [Fact]
        public void Should_change_font_of_the_column()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelColumn forthColumn = excelPackage.Workbook.Worksheets.First().Column(4);

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
        public void Should_set_horizontal_alignment_of_the_column()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelColumn forthColumn = excelPackage.Workbook.Worksheets.First().Column(4);

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
        public void Should_set_vertical_alignment_of_the_column()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelColumn forthColumn = excelPackage.Workbook.Worksheets.First().Column(4);

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
