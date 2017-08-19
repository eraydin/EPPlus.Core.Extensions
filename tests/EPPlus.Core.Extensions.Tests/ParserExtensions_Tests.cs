using FluentAssertions;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ParserExtensions_Tests : TestBase
    {
        [Fact]
        public void Test_TableValidation()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable table = excelPackage.GetTable("TEST3");
            List<ExcelTableConvertExceptionArgs> validation = table.Validate<WrongCars>().ToList();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------


            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            table.Should().NotBeNull("We have TEST3 table");
            validation.Should().NotBeNull("We have errors here");
            validation.Count.Should().Be(2, "We have 2 errors");

            validation.Exists(x => x.CellAddress.Address.Equals("C6", StringComparison.InvariantCultureIgnoreCase)).Should().BeTrue("Toyota is not in the enumeration");
            validation.Exists(x => x.CellAddress.Address.Equals("D7", StringComparison.InvariantCultureIgnoreCase)).Should().BeTrue("Date is null");
        }
    }

    enum Manufacturers { Opel = 1, Ford, Mercedes };
    class WrongCars
    {
        [ExcelTableColumn(ColumnName = "License plate")]
        public string LicensePlate { get; set; }

        [ExcelTableColumn]
        public Manufacturers Manufacturer { get; set; }

        [ExcelTableColumn(ColumnName = "Manufacturing date")]
        public DateTime ManufacturingDate { get; set; }

        [ExcelTableColumn(ColumnName = "Is ready for traffic?")]
        public bool Ready { get; set; }
    }
}
