using System;
using System.Drawing;
using System.Linq;

using EPPlus.Core.Extensions.Style;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using Xunit;

namespace EPPlus.Core.Extensions.Tests.StyleTests
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
            ExcelWorksheet firstWorksheet = workbook.GetWorksheet(1);
            ExcelWorksheet secondWorksheet = workbook.GetWorksheet(2);
            var nameOfStyle1 = "NamedStyle1";
            var nameOfStyle2 = "NamedStyle2";
            var fontName = "Arial";
            string valuedDimensionsOfSecondWorksheet = secondWorksheet.GetValuedDimension().Address;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------

            workbook.CreateNamedStyle(nameOfStyle1, style => style
                                                             .SetBackgroundColor(Color.Blue, ExcelFillStyle.DarkDown)
                                                             .SetFont(new Font(fontName, 12, FontStyle.Bold),
                                                                 Color.Yellow)
                                                             .BorderAround(ExcelBorderStyle.Double, Color.AliceBlue));

            workbook.CreateNamedStyleIfNotExists(nameOfStyle2, style => style
                                                                        .SetBackgroundColor(Color.Blue,
                                                                            ExcelFillStyle.DarkDown)
                                                                        .SetFont(new Font(fontName, 12, FontStyle.Bold),
                                                                            Color.Yellow)
                                                                        .BorderAround(ExcelBorderStyle.Double,
                                                                            Color.AliceBlue));

            firstWorksheet.Cells[1, 1, 1, 1].StyleName = nameOfStyle1;
            secondWorksheet.Cells[valuedDimensionsOfSecondWorksheet].StyleName = nameOfStyle1;
            secondWorksheet.Cells[valuedDimensionsOfSecondWorksheet].BorderAround(ExcelBorderStyle.DashDot);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            firstWorksheet.Cells[1, 1, 1, 1].StyleName.Should().Be(nameOfStyle1);
            firstWorksheet.Cells[1, 1, 1, 1].Style.Font.Name.Should().Be(fontName);
            firstWorksheet.Cells[1, 1, 1, 1].Style.Font.Color.Rgb.Should()
                          .Be(string.Format("{0:X8}", Color.Yellow.ToArgb() & 0xFFFFFFFF));
            firstWorksheet.Cells[1, 1, 1, 1].Style.Fill.BackgroundColor.Rgb.Should()
                          .Be(string.Format("{0:X8}", Color.Blue.ToArgb() & 0xFFFFFFFF));
            firstWorksheet.Cells[1, 1, 1, 1].Style.Fill.PatternType.Should().Be(ExcelFillStyle.DarkDown);
            firstWorksheet.Cells[1, 1, 1, 1].Style.Border.Top.Color.Rgb.Should()
                          .Be(string.Format("{0:X8}", Color.AliceBlue.ToArgb() & 0xFFFFFFFF));
            firstWorksheet.Cells[1, 1, 1, 1].Style.Border.Left.Style.Should().Be(ExcelBorderStyle.Double);

            secondWorksheet.Cells[valuedDimensionsOfSecondWorksheet].StyleName.Should().Be(nameOfStyle1);
            secondWorksheet.Cells[valuedDimensionsOfSecondWorksheet].Style.Font.Name.Should().Be(fontName);
            secondWorksheet.Cells[valuedDimensionsOfSecondWorksheet].Style.Font.Color.Rgb.Should()
                           .Be(string.Format("{0:X8}", Color.Yellow.ToArgb() & 0xFFFFFFFF));
            secondWorksheet.Cells[valuedDimensionsOfSecondWorksheet].Style.Fill.BackgroundColor.Rgb.Should()
                           .Be(string.Format("{0:X8}", Color.Blue.ToArgb() & 0xFFFFFFFF));
            secondWorksheet.Cells[valuedDimensionsOfSecondWorksheet].Style.Fill.PatternType.Should()
                           .Be(ExcelFillStyle.DarkDown);
            secondWorksheet.Cells[valuedDimensionsOfSecondWorksheet].Style.Border.Right.Style.Should()
                           .Be(ExcelBorderStyle.DashDot);
            secondWorksheet.Cells[valuedDimensionsOfSecondWorksheet].Style.Border.Bottom.Color.Rgb.Should()
                           .Be(string.Format("{0:X8}", Color.Black.ToArgb() & 0xFFFFFFFF));

            workbook.Styles.NamedStyles.Any(x => x.Name == nameOfStyle2).Should().Be(true);
        }

        [Fact]
        public void Should_not_create_again_already_defined_named_style()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorkbook workbook = excelPackage.Workbook;
            var nameOfStyle = "NamedStyle1";
            var fontName = "Arial";

            workbook.CreateNamedStyle(nameOfStyle, style =>
            {
                style.SetBackgroundColor(Color.Blue, ExcelFillStyle.DarkDown);
                style.SetFont(new Font(fontName, 12, FontStyle.Bold), Color.Yellow);
            });

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------

            Action action = () =>
            {
                workbook.CreateNamedStyle(nameOfStyle, style =>
                {
                    style.SetBackgroundColor(Color.Aquamarine, ExcelFillStyle.DarkGrid);
                    style.SetFont(new Font(fontName, 15, FontStyle.Italic), Color.Beige);
                });
            };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.Should().Throw<ArgumentException>();
        }

        [Fact]
        public void Should_not_throw_exception_for_already_defined_named_style()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorkbook workbook = excelPackage.Workbook;
            var nameOfStyle = "NamedStyle1";
            var fontName = "Arial";

            workbook.CreateNamedStyle(nameOfStyle, style =>
            {
                style.SetBackgroundColor(Color.Blue, ExcelFillStyle.DarkDown);
                style.SetFont(new Font(fontName, 12, FontStyle.Bold), Color.Yellow);
            });

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------

            Action action = () =>
            {
                workbook.CreateNamedStyleIfNotExists(nameOfStyle, style =>
                {
                    style.SetBackgroundColor(Color.Aquamarine, ExcelFillStyle.DarkGrid);
                    style.SetFont(new Font(fontName, 15, FontStyle.Italic), Color.Beige);
                });
            };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            workbook.Styles.NamedStyles.Count(x => x.Name == nameOfStyle).Should().Be(1);
            action.Should().NotThrow<Exception>();
        }
    }
}