using FluentAssertions;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ToExcelExtensions_Tests
    {
        [Fact]
        public void Test_ToExcelPackage()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            var cars = new List<Car>
                       {
                           new Car
                           {
                               Name = "Car1",
                               Price = 10m
                           },
                           new Car
                           {
                               Name = "Car2",
                               Price = 12.5m
                           }
                       };

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelPackage package = cars.ToExcelPackage("List of cars");

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            package.Workbook.Worksheets.First().Dimension.Rows.Should().Be(3);
        }
    }
}
