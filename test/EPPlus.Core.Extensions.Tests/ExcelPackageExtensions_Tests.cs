using System.Collections.Generic;
using System.Data;
using System.Linq;

using FluentAssertions;

using OfficeOpenXml.Table;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelPackageExtensions_Tests : TestBase
    {
        [Fact]
        public void Should_be_interceptable_current_row_when_its_converting_to_a_list()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            List<StocksNullable> stocksWithNullables;
            List<StocksValidation> stocksWithoutNullables;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            stocksWithNullables = excelPackage.ToList<StocksNullable>(6, configuration =>
                                                                         {
                                                                             configuration.SkipCastingErrors();
                                                                             configuration.SkipValidationErrors();
                                                                             configuration.WithHeaderRow();
                                                                             configuration.Intercept((current, rowIndex) => { current.Barcode = current.Barcode.Insert(0, "_"); });
                                                                         });

            stocksWithoutNullables = excelPackage.ToList<StocksValidation>(5, configuration =>
                                                                              {
                                                                                  configuration.SkipCastingErrors();
                                                                                  configuration.SkipValidationErrors();
                                                                                  configuration.WithHeaderRow();
                                                                                  configuration.Intercept((current, rowIndex) => { current.Quantity += 10 + rowIndex; });
                                                                              });

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            stocksWithNullables.Any().Should().BeTrue();
            stocksWithNullables.Count.Should().Be(3);
            stocksWithNullables.All(x => x.Barcode.StartsWith('_')).Should().Be(true);

            stocksWithoutNullables.Min(x => x.Quantity).Should().BeGreaterThan(10);
        }

        [Fact]
        public void Should_convert_an_Excel_package_into_a_DataSet()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            DataSet dataset;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            dataset = excelPackage.ToDataSet();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataset.Should().NotBeNull("We have 7 tables");
            dataset.Tables.Count.Should().Be(7, "We have 7 tables");
        }

        [Fact]
        public void Should_convert_given_Excel_package_to_list_of_objects()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            List<DateMap> list;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            list = excelPackage.ToList<DateMap>();

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
            List<StocksNullable> list;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            list = excelPackage.ToList<StocksNullable>(6, configuration =>
                                                          {
                                                              configuration.SkipCastingErrors();
                                                              configuration.WithoutHeaderRow();
                                                          });

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
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<ExcelTable> tables;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            tables = excelPackage.GetTables();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            tables.Should().NotBeNull("We have 4 tables");
            tables.Count().Should().Be(4, "We have 4 tables");

            excelPackage.HasTable("TEST2").Should().BeTrue("We have TEST2 table");
            excelPackage.HasTable("test2").Should().BeTrue("Table names are case insensitive");

            excelPackage.Workbook.Worksheets["TEST2"].Tables["TEST2"].Should().BeEquivalentTo(excelPackage.GetTable("TEST2"), "We are accessing the same objects");

            excelPackage.HasTable("NOTABLE").Should().BeFalse("We don't have NOTABLE table");
            excelPackage.GetTable("NOTABLE").Should().BeNull("We don't have NOTABLE table");
        }
    }
}
