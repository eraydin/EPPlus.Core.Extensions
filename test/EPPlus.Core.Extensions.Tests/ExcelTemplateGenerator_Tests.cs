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
            Type wrongCarsType = executingAssembly.FindExcelWorksheetTypes().First(x => x.Name == "WrongCars");

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage excelPackage1 = executingAssembly.GenerateExcelPackage(wrongCarsType.Name);

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
            Type wrongCarsType = executingAssembly.FindExcelWorksheetTypes().First(x => x.Name == "WrongCars");
            string defaultMapType = executingAssembly.GetNamesOfExcelWorksheetTypes().First(x => x == "DefaultMap");

            var excelPackage = new ExcelPackage();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.GenerateWorksheet(executingAssembly, wrongCarsType.Name);
            ExcelWorksheet worksheet2 = excelPackage.GenerateWorksheet(executingAssembly, defaultMapType);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            worksheet1.Should().NotBe(null);
            worksheet1.Name.Should().Be("Wrong Cars");
            worksheet1.GetColumns(1).Count().Should().BeGreaterThan(0);

            worksheet2.Should().NotBe(null);
            worksheet2.Name.Should().Be("DefaultMap");
            worksheet2.GetColumns(1).Count().Should().BeGreaterThan(0);
        }
    }
}
