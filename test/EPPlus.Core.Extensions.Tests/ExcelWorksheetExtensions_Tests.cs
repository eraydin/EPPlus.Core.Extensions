using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;

using EPPlus.Core.Extensions.Configuration;
using EPPlus.Core.Extensions.Validation;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelWorksheetExtensions_Tests : TestBase
    {
        [Fact]
        public void Test_GetDataBounds_With_Header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets["TEST5"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddress dataBounds = excelWorksheet.GetDataBounds();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataBounds.Rows.Should().Be(3);
            dataBounds.Columns.Should().Be(3);
        }

        [Fact]
        public void Test_GetDataBounds_Without_Header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets["TEST5"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddress dataBounds = excelWorksheet.GetDataBounds(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataBounds.Rows.Should().Be(4);
            dataBounds.Columns.Should().Be(3);
        }

        [Fact]
        public void Test_GetAsExcelTable_With_Header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets["TEST5"];
            IExcelConfiguration configuration = new ExcelConfiguration();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable excelTable = excelWorksheet.AsExcelTable();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------

            IList<StocksNullable> listOfStocks = excelTable.ToList<StocksNullable>(configuration);
            listOfStocks.Count.Should().Be(3);
        }

        [Fact]
        public void Test_GetAsExcelTable_Without_Header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets["TEST5"];
            IExcelConfiguration configuration = new ExcelConfiguration { SkipCastingErrors = true };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable excelTable = excelWorksheet.AsExcelTable(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------

            IList<StocksNullable> listOfStocks = excelTable.ToList<StocksNullable>(configuration);
            listOfStocks.Count.Should().Be(4);
        }

        [Fact]
        public void Test_ToDataTable_With_Headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            DataTable dataTable;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            dataTable = excelPackage.Workbook.Worksheets["TEST5"].ToDataTable();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataTable.Should().NotBeNull($"{nameof(dataTable)} should not be NULL");
            dataTable.Rows.Count.Should().Be(3, "We have 3 records");
        }

        [Fact]
        public void Test_ToDataTable_Without_Headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            DataTable dataTable;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            dataTable = excelPackage.Workbook.Worksheets["TEST5"].ToDataTable(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataTable.Should().NotBeNull($"{nameof(dataTable)} should not be NULL");
            dataTable.Rows.Count.Should().Be(4, "We have 4 records");
        }

        [Fact]
        public void Test_Worksheet_AsEnumerable()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.Workbook.Worksheets["TEST4"];
            ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets["TEST5"];
            IExcelConfiguration configuration1 = new ExcelConfiguration { SkipCastingErrors = true, HasHeaderRow = true };
            IExcelConfiguration configuration2 = new ExcelConfiguration { SkipCastingErrors = true, HasHeaderRow = false };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<StocksNullable> list1 = worksheet1.AsEnumerable<StocksNullable>(configuration1);
            IEnumerable<StocksNullable> list2 = worksheet2.AsEnumerable<StocksNullable>(configuration2);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list1.Count().Should().Be(4, "Should have four");
            list2.Count().Should().Be(4, "Should have four");
        }

        [Fact]
        public void Test_Worksheet_ToList()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.Workbook.Worksheets["TEST4"];
            ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets["TEST5"];
            IExcelConfiguration configuration1 = new ExcelConfiguration { SkipCastingErrors = true, HasHeaderRow = true };
            IExcelConfiguration configuration2 = new ExcelConfiguration { SkipCastingErrors = true, HasHeaderRow = false };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<StocksNullable> list1 = worksheet1.ToList<StocksNullable>(configuration1);
            IEnumerable<StocksNullable> list2 = worksheet2.ToList<StocksNullable>(configuration2);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list1.Count().Should().Be(4, "Should have four");
            list2.Count().Should().Be(4, "Should have four");
        }

        [Fact]
        public void Should_work_AddObjects_method_with_parameters()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TEST5"];
            IExcelConfiguration configuration1 = new ExcelConfiguration { SkipCastingErrors = false, HasHeaderRow = true };

            var stocks = new List<StocksNullable>
                         {
                             new StocksNullable
                             {
                                 Barcode = "barcode123",
                                 Quantity = 5,
                                 UpdatedDate = DateTime.MaxValue
                             }
                         };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddObjects(stocks, 5, _ => _.Barcode, _ => _.Quantity, _ => _.UpdatedDate);
            IEnumerable<StocksNullable> list = worksheet.ToList<StocksNullable>(configuration1);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count().Should().Be(4);
        }

        [Fact]
        public void Should_throw_exception_when_the_parameters_of_AddObjects_method_are_null()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TEST5"];
            var stocks = new List<StocksNullable>
                         {
                             new StocksNullable
                             {
                                 Barcode = "barcode123",
                                 Quantity = 5,
                                 UpdatedDate = DateTime.MaxValue
                             }
                         };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => { worksheet.AddObjects(stocks, 5, null); };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.ShouldThrow<ArgumentException>();
        }

        [Fact]
        public void Should_AddObjects_method_work_without_parameters()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TEST5"];
            IExcelConfiguration configuration = new ExcelConfiguration { SkipCastingErrors = true, HasHeaderRow = true };

            var stocks = new List<StocksNullable>
                         {
                             new StocksNullable
                             {
                                 Barcode = "barcode123",
                                 Quantity = 5,
                                 UpdatedDate = DateTime.MaxValue
                             }
                         };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddObjects(stocks, 5, 3);
            IEnumerable<StocksNullable> list = worksheet.ToList<StocksNullable>(configuration);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count().Should().Be(4);
        }

        [Fact]
        public void Should_AddLine_method_work()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TEST5"];
            IExcelConfiguration configuration = new ExcelConfiguration { SkipCastingErrors = false, HasHeaderRow = true };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddLine(5, "barcode123", 5, DateTime.UtcNow);
            IEnumerable<StocksNullable> list = worksheet.ToList<StocksNullable>(configuration);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count().Should().Be(4);
        }

        [Fact]
        public void Should_AddHeader_method_work()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TEST5"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddHeader(x =>
            {
                x.Style.Fill.PatternType = ExcelFillStyle.Solid;
                x.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(170, 170, 170));
            }, "Barcode", "Quantity", "UpdatedDate");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Dimension.End.Row.Should().Be(5);
        }

        [Fact]
        public void Should_HasAnyFormula_method_work()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.Workbook.Worksheets["TEST5"];
            worksheet1.Cells[18, 2, 18, 2].Formula = "=SUM(B2:B4)";

            ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets["TEST4"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool result1 = worksheet1.HasAnyFormula();
            bool result2 = worksheet2.HasAnyFormula();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            result1.Should().BeTrue();
            result2.Should().BeFalse();
        }

        [Fact]
        public void Test_GetColumns()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TEST6"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<KeyValuePair<int, string>> results = worksheet.GetColumns(1).ToList();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count().Should().Be(3);
            results.First().Value.Should().Be("Barcode");
        }

        [Fact]
        public void Test_CheckIfColumnValueIfNullOrEmpty()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TEST6"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool result1 = worksheet.CheckIfColumnValueIfNullOrEmpty(3, 4);
            bool result2 = worksheet.CheckIfColumnValueIfNullOrEmpty(2, 1);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            result1.Should().BeTrue();
            result2.Should().BeFalse();
        }

        [Fact]
        public void Test_StocksValidation()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TEST5"];
            IExcelConfiguration configuration = new ExcelConfiguration { SkipCastingErrors = false, HasHeaderRow = true };
            IList<StocksValidation> list;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => { list = worksheet.ToList<StocksValidation>(configuration); };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.ShouldThrow<ExcelTableValidationException>();
        }
    }
}
