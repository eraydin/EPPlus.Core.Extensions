using System;
using System.Collections.Generic;
using System.Linq;

using EPPlus.Core.Extensions.Configuration;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Table;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelTableExtensions_Tests : TestBase
    {
        /// <summary>
        ///     Test existence of test objects in the embedded workbook
        /// </summary>
        [Fact]
        public void WarmUp()
        {
            excelPackage.Should().NotBeNull("Excel package is null");

            // TEST1
            ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets["TEST1"];
            workSheet.Should().NotBeNull("Worksheet TEST1 missing");

            ExcelTable table = workSheet.Tables["TEST1"];
            table.Should().NotBeNull("Table TEST1 missing");

            table.Address.Columns.Should().Be(5, "Table1 is not as expected");
            table.Address.Rows.Should().BeGreaterThan(2, "Table1 has missing rows");

            // TEST2
            workSheet = excelPackage.Workbook.Worksheets["TEST2"];
            workSheet.Should().NotBeNull("Worksheet TEST2 missing");

            table = workSheet.Tables["TEST2"];
            table.Should().NotBeNull("Table TEST2 missing");

            table.Address.Columns.Should().Be(2, "Table2 is not as expected");
            table.Address.Rows.Should().BeGreaterThan(2, "Table2 has missing rows");
        }

        [Fact]
        public void Test_TableValidation()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.GetTable("TEST3");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<ExcelTableExceptionArgs> validation = table.Validate<WrongCars>().ToList();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            table.Should().NotBeNull("We have TEST3 table");
            validation.Should().NotBeNull("We have errors here");
            validation.Count.Should().Be(2, "We have 2 errors");

            validation.Exists(x => x.CellAddress.Address.Equals("C6", StringComparison.InvariantCultureIgnoreCase)).Should().BeTrue("Toyota is not in the enumeration");
            validation.Exists(x => x.CellAddress.Address.Equals("D7", StringComparison.InvariantCultureIgnoreCase)).Should().BeTrue("Date is null");
        }

        [Fact]
        public void Test_MapByDefault()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<DefaultMap> list = table.ToList<DefaultMap>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count.Should().Be(5, "We have expected 5 elements");
            list.First().Name.Should().Be("John", "We have expected John to be first");
            list.First().Gender.Should().Be("MALE", "We have expected a male to be first");
        }

        [Fact]
        public void Test_MapByName()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<NamedMap> list = table.ToList<NamedMap>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count.Should().Be(5, "We have expected 5 elements");
            list.First().FirstName.Should().Be("John", "We have expected John to be first");
            list.First().Sex.Should().Be("MALE", "We have expected a male to be first");
        }

        [Fact]
        public void Test_MapByIndex()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<IndexMap> list = table.ToList<IndexMap>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count.Should().Be(5, "We have expected 5 elements");
            list.First().Name.Should().Be("John", "We have expected John to be first");
            list.First().Gender.Should().Be("MALE", "We have expected a male to be first");
        }

        [Fact]
        public void Test_MapEnumString()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<EnumStringMap> list = table.ToList<EnumStringMap>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count.Should().Be(5, "We have expected 5 elements");
            list.Count(x => x.Gender == Genders.MALE).Should().Be(3, "We have expected 3 males");
            list.Count(x => x.Gender == Genders.FEMALE).Should().Be(2, "We have expected 2 females");
        }

        [Fact]
        public void Test_MapEnumNumeric()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<EnumByteMap> list = table.ToList<EnumByteMap>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count.Should().Be(5, "We have expected 5 elements");
            list.Count(x => x.Class == Classes.Ten).Should().Be(2, "We have expected 2 in 10th class");
            list.Count(x => x.Class == Classes.Nine).Should().Be(3, "We have expected 3 in 9th class");
        }

        /// <summary>
        ///     Test cases when a column is mapped to multiple properties (with different type)
        /// </summary>
        [Fact]
        public void Test_MultiMap()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<MultiMap> list = table.ToList<MultiMap>();
            MultiMap m = list.First(x => x.Class == Classes.Ten);
            MultiMap n = list.First(x => x.Class == Classes.Nine);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count.Should().Be(5, "We have expected 5 elements");

            ((int)m.Class).Should().Be(m.ClassAsInt, "Ten sould be 10");
            ((int)n.Class).Should().Be(n.ClassAsInt, "Nine sould be 9");
        }

        [Fact]
        public void Test_DateMap()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<DateMap> list = table.ToList<DateMap>();
            DateMap a = list.FirstOrDefault(x => x.Name == "Adam");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count.Should().Be(5, "We have expected 5 elements");

            a.BirthDate.Should().Be(new DateTime(1981, 4, 2), "Adam' birthday is 1981.04.02");

            list.Min(x => x.BirthDate).Should().Be(new DateTime(1979, 12, 1), "Oldest one was born on 1979.12.01");
        }

        [Fact]
        public void Test_MapFail()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];
            List<EnumFailMap> list;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => { list = table.ToList<EnumFailMap>(); };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.ShouldThrow<ExcelTableConvertException>()
                  .And.Args.CellValue.Should().Be("MALE");

            action.ShouldThrow<ExcelTableConvertException>()
                  .And.Args.ExpectedType.Should().Be(typeof(Classes));

            action.ShouldThrow<ExcelTableConvertException>()
                  .And.Args.PropertyName.Should().Be("Gender");

            action.ShouldThrow<ExcelTableConvertException>()
                  .And.Args.ColumnName.Should().Be("Gender");
        }

        [Fact]
        public void Test_MapSilentFail()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<EnumFailMap> list = table.ToList<EnumFailMap>(configuration =>
            {
                configuration.SkipCastingErrors = true;
            });

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Should().NotBeNull("We should get the list");

            list.All(x => !string.IsNullOrWhiteSpace(x.Name)).Should().BeTrue("All names should be there");
            list.All(x => x.Gender == 0).Should().BeTrue("All genders should be 0");
        }

        [Fact]
        public void Test_Nullable()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST2"].Tables["TEST2"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<CarNullable> list = table.ToList<CarNullable>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count(x => !x.Price.HasValue).Should().Be(2, "Should have two");
        }

        [Fact]
        public void Test_Nullable_With_SkipCastErrors()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST4"].Tables["TEST4"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<StocksNullable> list = table.ToList<StocksNullable>(configuration =>
            {
                configuration.SkipCastingErrors = true;
            });

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count(x => !x.UpdatedDate.HasValue).Should().Be(2, "Should have two");
            list.Count(x => !x.Quantity.HasValue).Should().Be(2, "Should have two");
            list.Count.Should().Be(4, "Should have four");
        }

        [Fact]
        public void Test_ComplexFixtures()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets["TEST3"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = workSheet.Tables["TEST3"];

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            workSheet.Should().NotBeNull("Worksheet TEST3 missing");
            table.Should().NotBeNull("Table TEST3 missing");

            table.Address.Columns.Should().Be(6, "Table3 is not as expected");
            table.Address.Rows.Should().Be(5 + (table.ShowTotal ? 1 : 0) + (table.ShowHeader ? 1 : 0), "Table3 has missing rows");
        }

        [Fact]
        public void Test_ComplexExample()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.GetTable("TEST3");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<Cars> list = table.ToList<Cars>();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count().Should().Be(5, "We have 5 rows");
            list.Count(x => string.IsNullOrWhiteSpace(x.LicensePlate)).Should().Be(1, "There is one without license plate");
            list.All(x => x.Manufacturer > 0).Should().BeTrue("All should have manufacturers");
            list.Last().ManufacturingDate.Should().BeNull("The last one's manufacturing date is unknown");
            list.Count(x => x.ManufacturingDate == null).Should().Be(1, "Only one manufacturig date is unknown");
            list.Single(x => x.LicensePlate == null).ShouldBeEquivalentTo(list.Single(x => !x.Ready), "The one without the license plate is not ready");
            list.Max(x => x.Price).Should().Be(12000, "Highest price is 12000");
            list.Max(x => x.ManufacturingDate).Should().Be(new DateTime(2015, 3, 10), "Oldest was manufactured on 2015.03.10");
        }
    }
}
