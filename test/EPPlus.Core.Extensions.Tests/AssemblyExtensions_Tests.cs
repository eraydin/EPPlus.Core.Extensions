using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using FluentAssertions;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class AssemblyExtensions_Tests
    {
        [Fact]
        public void Should_get_a_type_from_ExcelWorksheet_marked_types_by_name()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            KeyValuePair<string, string> wrongCars = executingAssembly.GetExcelWorksheetNamesOfMarkedTypes().First(x => x.Key == "WrongCars");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Type type = executingAssembly.GetExcelWorksheetMarkedTypeByName(wrongCars.Key);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            type.Should().NotBe(null);
            type.Name.Should().Be(wrongCars.Key);
        }

        [Fact]
        public void Should_get_all_ExcelWorksheet_marked_types()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<Type> results = executingAssembly.GetExcelWorksheetMarkedTypes();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Should_get_Excel_column_attributes_of_ExcelWorksheet_type()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            Type firstType = executingAssembly.GetExcelWorksheetMarkedTypes().First();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<ExcelTableColumnAttributeAndProperyInfo> results = firstType.GetExcelTableColumnAttributesWithProperyInfo();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Should_get_names_of_all_ExcelWorksheet_marked_objects()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<KeyValuePair<string, string>> results = executingAssembly.GetExcelWorksheetNamesOfMarkedTypes();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count.Should().BeGreaterThan(0);
            results.Any(x => x.Key.Equals("WrongCars") && x.Value.Equals("Wrong Cars")).Should().BeTrue();
            results.Any(x => x.Key.Equals("DefaultMap")).Should().BeTrue();
        }
    }
}
