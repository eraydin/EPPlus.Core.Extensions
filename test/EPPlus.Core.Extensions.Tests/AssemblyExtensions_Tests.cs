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
        public void Should_find_all_IExcelExportable_marked_types()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<Type> results = executingAssembly.FindExcelWorksheetTypes();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count.Should().BeGreaterThan(0);
        }

        [Fact]
        public void Should_get_a_type_from_name()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            string nameOfFirstType = executingAssembly.GetNamesOfExcelWorksheetTypes().First();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Type type = executingAssembly.FindExcelWorksheetByName(nameOfFirstType);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            type.Should().NotBe(null);
            type.Name.Should().Be(nameOfFirstType);
        }

        [Fact]
        public void Should_get_Excel_column_attributes_of_given_type()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            Type firstType = executingAssembly.FindExcelWorksheetTypes().First();

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
        public void Should_get_names_of_all_ExcelExportable_marked_objects()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<string> results = executingAssembly.GetNamesOfExcelWorksheetTypes();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            results.Count.Should().BeGreaterThan(0);
        }
    }
}
