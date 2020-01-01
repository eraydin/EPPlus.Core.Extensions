using System.Collections.Generic;
using System.Data;
using System.Linq;

using FluentAssertions;

using OfficeOpenXml;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelPackageExtensionsTests : TestBase
    {
        [Fact]
        public void Should_add_a_copied_worksheet_to_the_package()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            string randomName = GetRandomName();
            ExcelWorksheet copyWorksheet = ExcelPackage1.GetWorksheet(2);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage1.AddWorksheet(randomName, copyWorksheet);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage1.GetWorksheet(randomName).Should().NotBe(null);
            ExcelPackage1.GetWorksheet(randomName).Cells[1, 1, 3, 3].Value.Should().BeEquivalentTo(copyWorksheet.Cells[1, 1, 3, 3].Value);
        }

        [Fact]
        public void Should_add_empty_worksheet_to_the_package()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            string randomName = GetRandomName();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage1.AddWorksheet(randomName);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage1.GetWorksheet(randomName).Should().NotBe(null);
        }

        [Fact]
        public void Should_be_interceptable_current_row_when_its_converting_to_a_list()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            var stocksNullableWorksheetIndex = 5;
            var stocksValidationWorksheetIndex = 4;

#if NETFRAMEWORK
            stocksNullableWorksheetIndex = 6;
            stocksValidationWorksheetIndex = 5;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            var optionalStocks = ExcelPackage1.ToList<StocksNullable>(stocksNullableWorksheetIndex, configuration => configuration.SkipCastingErrors()
                                                                                                                                       .SkipValidationErrors()
                                                                                                                                       .Intercept((current, rowIndex) => { current.Barcode = current.Barcode.Insert(0, "_"); }));

            var stocks = ExcelPackage1.ToList<StocksValidation>(stocksValidationWorksheetIndex, configuration => configuration.SkipCastingErrors()
                                                                                                                                              .SkipValidationErrors()
                                                                                                                                              .Intercept((current, rowIndex) => { current.Quantity += 10 + rowIndex; }));

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            optionalStocks.Any().Should().BeTrue();
            optionalStocks.Count.Should().Be(3);
            optionalStocks.All(x => x.Barcode.StartsWith("_")).Should().Be(true);

            stocks.Min(x => x.Quantity).Should().BeGreaterThan(10);
        }

        [Fact]
        public void Should_convert_an_Excel_package_into_a_DataSet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            DataSet dataset;
            const int expectedCount = 8;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            dataset = ExcelPackage1.ToDataSet();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataset.Should().NotBeNull($"We have {expectedCount} tables");
            dataset.Tables.Count.Should().Be(expectedCount, $"We have {expectedCount} tables");
        }

        [Fact]
        public void Should_convert_given_Excel_package_to_list_of_objects()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            var list = ExcelPackage1.ToList<DateMap>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Any().Should().BeTrue();
            list.Count.Should().Be(5);
        }

        [Fact]
        public void Should_convert_given_ExcelPackage_to_list_of_objects_with_worksheet_index()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            var worksheetIndex = 5;
#if NETFRAMEWORK
            worksheetIndex = 6;
#endif

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            var list = ExcelPackage1.ToList<StocksNullable>(worksheetIndex, configuration => configuration.SkipCastingErrors()
                                                                                                          .WithoutHeaderRow());

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Any().Should().BeTrue();
            list.Count.Should().Be(4);
        }

        [Fact]
        public void Should_extract_all_tables_from_an_Excel_package()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            var tables = ExcelPackage1.GetAllTables().ToList();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            tables.Should().NotBeNull("We have 4 tables");
            tables.Count().Should().Be(4, "We have 4 tables");

            ExcelPackage1.HasTable("TEST2").Should().BeTrue("We have TEST2 table");
            ExcelPackage1.HasTable("test2").Should().BeTrue("Table names are case insensitive");

            ExcelPackage1.GetWorksheet("TEST2").GetTable("TEST2").Should().BeEquivalentTo(ExcelPackage1.GetTable("TEST2"), "We are accessing the same objects");

            ExcelPackage1.HasTable("NOTABLE").Should().BeFalse("We don't have NOTABLE table");
            ExcelPackage1.GetTable("NOTABLE").Should().BeNull("We don't have NOTABLE table");
        }
    }
}
