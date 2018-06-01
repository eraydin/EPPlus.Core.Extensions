using System;
using System.Linq;
using System.Reflection;

using FluentAssertions;

using OfficeOpenXml;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelTemplateGenerator_Tests
    {
        [Fact]
        public void Should_generate_an_Excel_package_from_given_ExcelExportable_class_name()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            Type type = executingAssembly.FindExcelExportableTypes().First();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage excelPackage1 = executingAssembly.GenerateExcelPackage(type.Name);

            Action act = () => executingAssembly.GenerateExcelPackage("sadas");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            excelPackage1.Should().NotBe(null);
            excelPackage1.GetWorksheet(1).GetColumns(1).Count().Should().BeGreaterThan(0);

            act.Should().Throw<ArgumentNullException>().And.ParamName.Should().Be("typeName");
        }

        [Fact]
        public void Should_generate_an_worksheet_from_given_ExcelExportable_class_name()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            Type firstType = executingAssembly.FindExcelExportableTypes().First();

            var excelPackage = new ExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.GenerateWorksheet(executingAssembly, firstType.Name);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet1.Should().NotBe(null);
            worksheet1.GetColumns(1).Count().Should().BeGreaterThan(0);
        }
    }
}
