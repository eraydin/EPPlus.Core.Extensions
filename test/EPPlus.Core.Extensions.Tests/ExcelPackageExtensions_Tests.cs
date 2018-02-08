using System.Collections.Generic;
using System.Data;
using System.Linq;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Table;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelPackageExtensions_Tests : TestBase
    {
        [Fact]
        public void Should_extract_all_excelTables_from_an_excelPackage()
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

            excelPackage.Workbook.Worksheets["TEST2"].Tables["TEST2"].ShouldBeEquivalentTo(excelPackage.GetTable("TEST2"), "We are accessing the same objects");

            excelPackage.HasTable("NOTABLE").Should().BeFalse("We don't have NOTABLE table");
            excelPackage.GetTable("NOTABLE").Should().BeNull("We don't have NOTABLE table");
        }

        [Fact]
        public void Should_convert_an_excelPackage_into_a_dataSet()
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
            dataset.Should().NotBeNull("We have 6 tables");
            dataset.Tables.Count.Should().Be(6, "We have 6 tables");
        }

        [Fact]
        public void Should_convert_a_byte_array_into_an_excelPackage()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            byte[] buffer = excelPackage.GetAsByteArray();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package = buffer.GetAsExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.Should().NotBeNull();
            package.Workbook.Worksheets.Count.ShouldBeEquivalentTo(excelPackage.Workbook.Worksheets.Count);
            package.GetTables().Count().ShouldBeEquivalentTo(excelPackage.GetTables().Count());
        }

        [Fact]
        public void Should_convert_given_ExcelPackage_to_list_of_objects()
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
                configuration.HasHeaderRow = false;
                configuration.SkipCastingErrors = true;
            });

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Any().Should().BeTrue();
            list.Count.Should().Be(4);
        }
    }
}
