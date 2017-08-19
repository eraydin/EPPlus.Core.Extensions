using FluentAssertions;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelPackageExtensions_Tests : TestBase
    {
        [Fact]
        public void Test_TableNameExtensions()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<ExcelTable> tables = excelPackage.GetTables();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------


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
    }
}
