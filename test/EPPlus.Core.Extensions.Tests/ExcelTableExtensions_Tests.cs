using System;
using System.Collections.Generic;
using System.Linq;

using EPPlus.Core.Extensions.Exceptions;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Table;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelTableExtensions_Tests : TestBase
    {
        [Fact]
        public void Should_get_databounds_of_Excel_table_including_header_row()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable tableWithoutHeaderRow = excelPackage.Workbook.Worksheets["TEST6"].AsExcelTable(false);  

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddress dataBoundsIncludingHeader = tableWithoutHeaderRow.GetDataBounds();
              
            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //----------------------------------------------------------------------------------------------------------- 
            dataBoundsIncludingHeader.Start.Row.Should().Be(2);
            dataBoundsIncludingHeader.End.Row.Should().Be(5);
            dataBoundsIncludingHeader.Address.Should().Be("A2:C5");
        }

        [Fact]
        public void Should_get_databounds_of_Excel_table_without_header_row()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable tableWithHeaderRow = excelPackage.Workbook.Worksheets["TEST6"].AsExcelTable();        

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddress dataBoundsWithoutHeader = tableWithHeaderRow.GetDataBounds();
                                  
            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataBoundsWithoutHeader.Start.Row.Should().Be(3);
            dataBoundsWithoutHeader.End.Row.Should().Be(5);
            dataBoundsWithoutHeader.Address.Should().Be("A3:C5");
        }

        [Fact]
        public void Should_get_ExcelTable_as_enumerable()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable tableWithoutHeaderRow = excelPackage.Workbook.Worksheets["TEST6"].AsExcelTable(false);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<StocksNullable> result = tableWithoutHeaderRow.AsEnumerable<StocksNullable>(c => c.SkipCastingErrors());

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //----------------------------------------------------------------------------------------------------------- 
            result.Count().Should().Be(4);
        }

        [Fact]
        public void Should_throw_an_exception_when_trying_to_get_an_excel_table_if_casting_error_occurred()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable tableWithoutHeaderRow = excelPackage.Workbook.Worksheets["TEST6"].AsExcelTable(false);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------    
            Action act = () =>
                         {
                             List<StocksNullable> result = tableWithoutHeaderRow.AsEnumerable<StocksNullable>(c => c.WithCastingExceptionMessage("Casting error occured on '{2}'")).ToList();
                         };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //----------------------------------------------------------------------------------------------------------- 
            act.Should().Throw<ExcelException>().WithMessage("Casting error occured on 'B2'"); 
        }     

        [Fact]
        public void Should_automatically_map_if_column_name_and_index_not_specified()
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
        public void Should_convert_Excel_table_into_list_of_complex_objects()
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
            list.Count.Should().Be(5, "We have 5 rows");

            // 'LicensePlate' property should be mapped by column index
            list.Count(x => string.IsNullOrWhiteSpace(x.LicensePlate)).Should().Be(1, "There is one without license plate");

            // 'Manufacturer' property should be mapped to Manufacturers2 enum
            list.All(x => x.Manufacturer > 0).Should().BeTrue("All should have manufacturers");

            // 'ManufacturingDate' property should be mapped by column name "Manufacturing date"
            list.Last().ManufacturingDate.Should().BeNull("The last one's manufacturing date is unknown");
            list.Count(x => x.ManufacturingDate == null).Should().Be(1, "Only one manufacturig date is unknown");
            list.Max(x => x.ManufacturingDate).Should().Be(new DateTime(2015, 3, 10), "Oldest was manufactured on 2015.03.10");

            // 'Ready' property should be mapped to "Is ready for traffic?" column
            list.Single(x => x.LicensePlate == null).Should().BeEquivalentTo(list.Single(x => !x.Ready), "The one without the license plate is not ready");

            // 'Price' property should be mapped automatically
            list.Max(x => x.Price).Should().Be(12000, "Highest price is 12000");
        }

        [Fact]
        public void Should_convert_Excel_table_into_list_of_nullable_objects()
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
            list.Count.Should().Be(5);
        }

        [Fact]
        public void Should_convert_Excel_table_into_list_of_objects()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<DateMap> list = table.ToList<DateMap>();   

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count.Should().Be(5, "We have expected 5 elements");

            list.First(x => x.Name == "Adam").BirthDate.Should().Be(new DateTime(1981, 4, 2), "Adam' birthday is 1981.04.02");

            list.Min(x => x.BirthDate).Should().Be(new DateTime(1979, 12, 1), "Oldest one was born on 1979.12.01");
        }

        /// <summary>
        ///     Test cases when a column is mapped to multiple properties (with different type)
        /// </summary>
        [Fact]
        public void Should_map_a_column_into_multiple_properties_of_an_object()
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
        public void Should_map_integer_values_into_Enum()
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

        [Fact]
        public void Should_map_object_properties_with_column_index()
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
        public void Should_map_object_properties_with_different_column_names()
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
            list.First().NotMapped.Should().Be(null);
        }

        [Fact]
        public void Should_map_string_values_into_Enums()
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
        public void Should_not_throw_exception_if_casting_error_occured_with_nullable()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST4"].Tables["TEST4"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<StocksNullable> list = table.ToList<StocksNullable>(configuration => configuration.SkipCastingErrors());

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Count(x => !x.UpdatedDate.HasValue).Should().Be(2, "Should have two");
            list.Count(x => !x.Quantity.HasValue).Should().Be(2, "Should have two");
            list.Count.Should().Be(4, "Should have four");
        }

        [Fact]
        public void Should_not_throw_exception_if_casting_failed()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<EnumFailMap> list = table.ToList<EnumFailMap>(configuration => configuration.SkipCastingErrors());

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list.Should().NotBeNull("We should get the list");

            list.All(x => !string.IsNullOrWhiteSpace(x.Name)).Should().BeTrue("All names should be there");
            list.All(x => x.Gender == 0).Should().BeTrue("All genders should be 0");
        }

        [Fact]
        public void Should_throw_exception_if_casting_failed()
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
            action.Should().Throw<ExcelException>()
                  .And.Args.CellValue.Should().Be("MALE");

            action.Should().Throw<ExcelException>()
                  .And.Args.ExpectedType.Should().Be(typeof(Classes));

            action.Should().Throw<ExcelException>()
                  .And.Args.PropertyName.Should().Be("Gender");

            action.Should().Throw<ExcelException>()
                  .And.Args.ColumnName.Should().Be("Gender");
        }

        [Fact]
        public void Should_throw_argument_exception_if_there_is_no_mapped_property_object()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];
            List<ObjectWithoutExcelTableAttributes> list;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => { list = table.ToList<ObjectWithoutExcelTableAttributes>(); };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.Should().Throw<ArgumentException>().WithMessage($"Given object does not have any {nameof(ExcelTableColumnAttribute)}.");   
        }

        //[Fact]
        public void Should_throw_exception_if_object_properties_not_mapped_correctly()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];
            List<ObjectWithWrongAttributeMappings> list;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action = () => { list = table.ToList<ObjectWithWrongAttributeMappings>(); };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action.Should().Throw<ArgumentException>().WithMessage($"Given object does not have any {nameof(ExcelTableColumnAttribute)}.");
        }

        [Fact]
        public void Should_validate_Excel_table_and_return_casting_errors()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.GetTable("TEST3");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<ExcelExceptionArgs> validationResults = table.Validate<WrongCars>().ToList();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            table.Should().NotBeNull("We have TEST3 table");
            validationResults.Should().NotBeNull("We have errors here");
            validationResults.Count.Should().Be(2, "We have 2 errors");

            validationResults.Exists(x => x.CellAddress.Address.Equals("C6", StringComparison.InvariantCultureIgnoreCase)).Should().BeTrue("Toyota is not in the enumeration");
            validationResults.Exists(x => x.CellAddress.Address.Equals("D7", StringComparison.InvariantCultureIgnoreCase)).Should().BeTrue("Date is null");
        }
    }
}
