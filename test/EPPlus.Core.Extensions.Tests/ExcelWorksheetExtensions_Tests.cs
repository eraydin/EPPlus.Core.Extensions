using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;

using EPPlus.Core.Extensions.Exceptions;
using EPPlus.Core.Extensions.Style;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Table;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelWorksheetExtensions_Tests : TestBase
    {
        [Fact]
        public void Should_add_an_header_without_configuration()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddHeader("NewBarcode", "NewQuantity", "NewUpdatedDate");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Dimension.End.Row.Should().Be(5);
            worksheet.Cells[1, 1, 1, 1].Value.Should().Be("NewBarcode");
            worksheet.Cells[1, 2, 1, 2].Value.Should().Be("NewQuantity");
            worksheet.Cells[1, 3, 1, 3].Value.Should().Be("NewUpdatedDate");
            worksheet.Cells[2, 1, 2, 1].Value.Should().Be("Barcode");
        }

        [Fact]
        public void Should_add_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST5");
            Color color = Color.AntiqueWhite;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddHeader(x => x.SetBackgroundColor(color), "NewBarcode", "NewQuantity", "NewUpdatedDate");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Dimension.End.Row.Should().Be(5);
            worksheet.Cells[1, 1, 1, 1].Value.Should().Be("NewBarcode");
            worksheet.Cells[1, 2, 1, 2].Value.Should().Be("NewQuantity");
            worksheet.Cells[1, 3, 1, 3].Value.Should().Be("NewUpdatedDate");

            worksheet.Cells[1, 1, 1, 1].Style.Fill.BackgroundColor.Rgb.Should()
                     .Be(string.Format("{0:X8}", color.ToArgb() & 0xFFFFFFFF));
            worksheet.Cells[2, 1, 2, 1].Value.Should().Be("Barcode");
        }

        [Fact]
        public void Should_add_line_with_configuration()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddLine(5, configureCells => configureCells.SetBackgroundColor(Color.Yellow), "barcode123", 5,
                DateTime.UtcNow);
            IEnumerable<StocksNullable> list = worksheet.ToList<StocksNullable>(configuration => configuration
                                                                                                 .WithoutHeaderRow()
                                                                                                 .SkipCastingErrors());

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.Cells[5, 1].Style.Fill.BackgroundColor.Rgb.Should()
                     .Be(string.Format("{0:X8}", Color.Yellow.ToArgb() & 0xFFFFFFFF));
            list.Count().Should().Be(5);
        }

        [Fact]
        public void Should_add_line_without_configuration()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddLine(5, "barcode123", 5, DateTime.UtcNow);
            IEnumerable<StocksNullable> list = worksheet.ToList<StocksNullable>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count().Should().Be(4);
        }

        [Fact]
        public void Should_add_objects_with_parameters()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST5");
            DateTime dateTime = DateTime.MaxValue;

            var stocks = new List<StocksNullable>
            {
                new StocksNullable
                {
                    Barcode = "barcode123",
                    Quantity = 5,
                    UpdatedDate = dateTime
                }
            };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.AddObjects(stocks, 5, _ => _.Barcode, _ => _.Quantity, _ => _.UpdatedDate);
            IEnumerable<StocksNullable> list = worksheet.ToList<StocksNullable>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count().Should().Be(4);
            list.Last().Barcode.Should().Be("barcode123");
            list.Last().Quantity.Should().Be(5);
            list.Last().UpdatedDate.Value.Date.Should().Be(dateTime.Date);
            list.Last().UpdatedDate.Value.Hour.Should().Be(dateTime.Hour);
        }

        [Fact]
        public void Should_add_objects_with_start_row_and_column_index_without_parameters()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST5");

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
            IEnumerable<StocksNullable> list =
                worksheet.ToList<StocksNullable>(configuration => configuration.SkipCastingErrors());

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count().Should().Be(4);
        }

        [Fact]
        public void Should_cannot_add_objects_with_null_property_selectors()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST5");
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
            action.Should().Throw<ArgumentException>();
        }

        [Fact]
        public void Should_check_and_throw_exception_if_column_value_is_wrong_on_specified_index()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action1 = () => { worksheet.CheckAndThrowColumn(2, 3, "Barcode", "Barcode column is missing"); };

            Action action2 = () => { worksheet.CheckAndThrowColumn(2, 1, "Barcode"); };

            Action action3 = () => { worksheet.CheckAndThrowColumn(3, 14, "Barcode"); };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action1.Should().Throw<ExcelValidationException>().And.Message.Should().Be("Barcode column is missing");
            action2.Should().NotThrow();
            action3.Should().Throw<ExcelValidationException>();
        }

        [Fact]
        public void Should_check_if_column_value_is_null_or_empty_on_given_index()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool result1 = worksheet.CheckColumnValueIsNullOrEmpty(3, 4);
            bool result2 = worksheet.CheckColumnValueIsNullOrEmpty(2, 1);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            result1.Should().BeTrue();
            result2.Should().BeFalse();
        }

        [Fact]
        public void Should_convert_to_datatable_with_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            DataTable dataTable;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            dataTable = excelPackage.GetWorksheet("TEST5").ToDataTable();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataTable.Should().NotBeNull($"{nameof(dataTable)} should not be NULL");
            dataTable.Rows.Count.Should().Be(3, "We have 3 records");
        }

        [Fact]
        public void Should_convert_to_datatable_without_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            DataTable dataTable;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            dataTable = excelPackage.GetWorksheet("TEST5").ToDataTable(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataTable.Should().NotBeNull($"{nameof(dataTable)} should not be NULL");
            dataTable.Rows.Count.Should().Be(4, "We have 4 records");
        }

        [Fact]
        public void Should_convert_worksheet_to_list()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.GetWorksheet("TEST4");
            ExcelWorksheet worksheet2 = excelPackage.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<StocksNullable> list1 = worksheet1.ToList<StocksNullable>(configuration =>
            {
                configuration.SkipCastingErrors();
                configuration.WithoutHeaderRow();
            });

            List<StocksNullable> list2 = worksheet2.ToList<StocksNullable>(configuration =>
            {
                configuration.SkipCastingErrors();
                configuration.WithoutHeaderRow();
            });

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list1.Count().Should().Be(4, "Should have four");
            list2.Count().Should().Be(4, "Should have four");
        }

        [Fact]
        public void Should_delete_a_column_by_using_header_text()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.DeleteColumn("Quantity");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.GetValuedDimension().End.Column.Should().Be(2);
            worksheet.Cells[2, 2, 2, 2].Text.Should().Be("UpdatedDate");
        }

        [Fact]
        public void Should_delete_columns_by_given_header_text()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST6");

            ExcelAddressBase valuedDimension = worksheet.GetValuedDimension();

            worksheet.ChangeCellValue(2, valuedDimension.End.Column + 1, "Quantity");
            worksheet.ChangeCellValue(2, valuedDimension.End.Column + 2, "Quantity");
            worksheet.ChangeCellValue(2, valuedDimension.End.Column + 3, "Quantity");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            worksheet.DeleteColumns("Quantity");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet.GetValuedDimension().End.Column.Should().Be(2);
            worksheet.Cells[2, 2, 2, 2].Text.Should().Be("UpdatedDate");
        }

        [Fact]
        public void Should_found_any_formula_on_worksheet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.GetWorksheet("TEST5");
            worksheet1.Cells[18, 2, 18, 2].Formula = "=SUM(B2:B4)";

            ExcelWorksheet worksheet2 = excelPackage.GetWorksheet("TEST4");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool result1 = worksheet1.HasAnyFormula();
            bool result2 = worksheet2.HasAnyFormula();
            Action action1 = () => { worksheet1.CheckAndThrowIfThereIsAnyFormula("First worksheet has formulas."); };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            result1.Should().BeTrue();
            result2.Should().BeFalse();
            action1.Should().Throw<ExcelValidationException>().WithMessage("First worksheet has formulas.");
        }

        [Fact]
        public void Should_get_as_Excel_table_with_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable excelTable = excelWorksheet.AsExcelTable();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------

            List<StocksNullable> listOfStocks = excelTable.ToList<StocksNullable>();
            listOfStocks.Count.Should().Be(3);
        }

        [Fact]
        public void Should_get_as_Excel_table_without_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable excelTable = excelWorksheet.AsExcelTable(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------

            List<StocksNullable> listOfStocks =
                excelTable.ToList<StocksNullable>(configuration => configuration.SkipCastingErrors());
            listOfStocks.Count.Should().Be(4);
        }

        [Fact]
        public void Should_get_columns_of_given_row_index()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<KeyValuePair<int, string>> results = worksheet.GetColumns(2).ToList();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count().Should().Be(3);
            results.First().Value.Should().Be("Barcode");
        }

        [Fact]
        public void Should_get_data_bounds_of_worksheet_with_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.GetWorksheet("TEST5");

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
        public void Should_get_data_bounds_of_worksheet_without_headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.GetWorksheet("TEST5");

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
        public void Should_get_empty_list_if_table_is_empty()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TEST7"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<StocksNullable> results = worksheet.ToList<StocksNullable>(configuration =>
                configuration.Intercept((item, row) => { item.Barcode = item.Barcode.Trim(); })
            );

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count.Should().Be(0);
        }

        [Fact]
        public void Should_get_worksheet_as_enumerable()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.GetWorksheet("TEST4");
            ExcelWorksheet worksheet2 = excelPackage.GetWorksheet("TEST5");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<StocksNullable> list1 = worksheet1.AsEnumerable<StocksNullable>(configuration =>
            {
                configuration.SkipCastingErrors();
                configuration.SkipValidationErrors();
                configuration.WithoutHeaderRow();
            });

            IEnumerable<StocksNullable> list2 = worksheet2.AsEnumerable<StocksNullable>(configuration =>
            {
                configuration.SkipCastingErrors();
                configuration.SkipValidationErrors();
                configuration.WithoutHeaderRow();
            });

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list1.Count().Should().Be(4, "Should have four");
            list2.Count().Should().Be(4, "Should have four");
        }

        [Fact]
        public void Should_parse_datetime_value_as_correctly_if_formatted_customly()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<StocksNullable> nullableStocks = worksheet.ToList<StocksNullable>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            nullableStocks[0].UpdatedDate.HasValue.Should().Be(true);
            nullableStocks[0].UpdatedDate.Value.Date.Should().Be(new DateTime(2017, 08, 08));

            nullableStocks[1].UpdatedDate.HasValue.Should().Be(true);
            nullableStocks[1].UpdatedDate.Value.Should().Be(new DateTime(2016, 11, 03, 01, 30, 53));
        }

        [Fact]
        public void Should_throw_an_exception_when_columns_of_worksheet_not_matched_with_ExcelTableAttribute()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.GetWorksheet("TEST6");
            ExcelWorksheet emptySheet1 = excelPackage.GetWorksheet("EmptySheet");
            ExcelWorksheet emptySheet2 = excelPackage.GetWorksheet("EmptySheet");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action1 = () =>
            {
                worksheet1.CheckHeadersAndThrow<NamedMap>(2, "The {0}.column of worksheet should be '{1}'.");
            };
            Action action2 = () => { worksheet1.CheckHeadersAndThrow<NamedMap>(2); };
            Action action3 = () => { worksheet1.CheckHeadersAndThrow<StocksNullable>(2); };
            Action action4 = () => { worksheet1.CheckHeadersAndThrow<Car>(2); };

            Action actionForEmptySheet1 = () =>
                emptySheet1.CheckHeadersAndThrow<StocksValidation>(1, "The {0}.column of worksheet should be '{1}'.");
            Action actionForEmptySheet2 = () =>
                emptySheet2.CheckHeadersAndThrow<Cars>(1, "The {0}.column of worksheet should be '{1}'.");


            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action1.Should().Throw<ExcelValidationException>().And.Message.Should()
                   .Be("The 1.column of worksheet should be 'Name'.");
            action2.Should().Throw<ExcelValidationException>().And.Message.Should()
                   .Be("The 1. column of worksheet should be 'Name'.");
            action3.Should().NotThrow<ExcelValidationException>();
            action4.Should().Throw<ArgumentException>();

            actionForEmptySheet1.Should().Throw<ExcelValidationException>().And.Message.Should()
                                .Be("The 1.column of worksheet should be 'Barcode'.");
            actionForEmptySheet2.Should().Throw<ExcelValidationException>().And.Message.Should()
                                .Be("The 1.column of worksheet should be 'LicensePlate'.");
        }

        [Fact]
        public void Should_throw_Excel_validation_exception_if_worksheet_does_not_have_valued_dimension()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("EmptySheet");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<Cars> cars = worksheet.ToList<Cars>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            cars.Count.Should().Be(0);
        }

        [Fact]
        public void Should_throw_exception_when_occured_validation_error()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST5");
            List<StocksValidation> list;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => { list = worksheet.ToList<StocksValidation>(); };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.Should().Throw<ExcelValidationException>()
                  .WithMessage("Please enter a value bigger than 10")
                  .And.Args.ColumnName.Should().Be("Quantity");
        }

        [Fact]
        public void Should_valued_dimension_be_A2C5()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST6");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddressBase valuedDimension = worksheet.GetValuedDimension();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            valuedDimension.Address.Should().Be("A2:C5");
            valuedDimension.Start.Column.Should().Be(1);
            valuedDimension.Start.Row.Should().Be(2);
            valuedDimension.End.Column.Should().Be(3);
            valuedDimension.End.Row.Should().Be(5);
        }

        [Fact]
        public void Should_valued_dimension_be_E9G13()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet = excelPackage.GetWorksheet("TEST4");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddressBase valuedDimension = worksheet.GetValuedDimension();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            valuedDimension.Address.Should().Be("E9:G13");
            valuedDimension.Start.Column.Should().Be(5);
            valuedDimension.Start.Row.Should().Be(9);
            valuedDimension.End.Column.Should().Be(7);
            valuedDimension.End.Row.Should().Be(13);
        }
    }
}