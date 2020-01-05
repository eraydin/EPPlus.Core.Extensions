using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using EPPlus.Core.Extensions.Attributes;

using FluentAssertions;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class AssemblyExtensionsTests
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
            List<Type> results = executingAssembly.GetTypesMarkedAsExcelWorksheet();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Should().HaveCountGreaterThan(0);
        }

        [Fact]
        public void Should_get_Excel_column_attributes_of_ExcelWorksheet_type()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            Type firstType = executingAssembly.GetTypesMarkedAsExcelWorksheet().First();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<ExcelTableColumnDetails> results = firstType.GetExcelTableColumnAttributesWithPropertyInfo();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Should().HaveCountGreaterThan(0);
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
            results.Should().HaveCountGreaterThan(0);
            results.Should().Contain(x => x.Key.Equals("WrongCars") && x.Value.Equals("Wrong Cars"));
            results.Should().Contain(x => x.Key.Equals("DefaultMap"));
        }
    }
}
