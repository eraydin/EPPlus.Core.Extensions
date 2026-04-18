using System.Collections.Generic;
using System.Linq;

using EPPlus.Core.Extensions.Attributes;

using FluentAssertions;

using OfficeOpenXml;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class HeaderRowIndexTests
    {
        [Fact]
        public void ToList_should_read_data_when_header_is_not_on_row_1()
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPlus.Core.Extensions");
            using var package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

            // Row 1: report title (should be skipped)
            ws.Cells[1, 1].Value = "Monthly Report";

            // Row 2: actual header
            ws.Cells[2, 1].Value = "Name";
            ws.Cells[2, 2].Value = "Gender";

            // Rows 3-4: data
            ws.Cells[3, 1].Value = "Alice";
            ws.Cells[3, 2].Value = "Female";
            ws.Cells[4, 1].Value = "Bob";
            ws.Cells[4, 2].Value = "Male";

            List<NamedMap> result = ws.ToList<NamedMap>(c => c.WithHeaderRowIndex(2));

            result.Should().HaveCount(2);
            result[0].FirstName.Should().Be("Alice");
            result[1].FirstName.Should().Be("Bob");
        }

        [Fact]
        public void ToList_without_HeaderRowIndex_should_still_work_as_before()
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPlus.Core.Extensions");
            using var package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

            ws.Cells[1, 1].Value = "Name";
            ws.Cells[1, 2].Value = "Gender";
            ws.Cells[2, 1].Value = "Alice";
            ws.Cells[2, 2].Value = "Female";

            List<NamedMap> result = ws.ToList<NamedMap>();

            result.Should().HaveCount(1);
            result[0].FirstName.Should().Be("Alice");
        }
    }

    public class NestedColumnTests
    {
        [Fact]
        public void ToList_should_map_nested_column_properties()
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPlus.Core.Extensions");
            using var package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

            ws.Cells[1, 1].Value = "Name";
            ws.Cells[1, 2].Value = "Street";
            ws.Cells[1, 3].Value = "City";

            ws.Cells[2, 1].Value = "Alice";
            ws.Cells[2, 2].Value = "123 Main St";
            ws.Cells[2, 3].Value = "London";

            ws.Cells[3, 1].Value = "Bob";
            ws.Cells[3, 2].Value = "456 Elm St";
            ws.Cells[3, 3].Value = "Paris";

            List<PersonWithAddress> result = ws.ToList<PersonWithAddress>();

            result.Should().HaveCount(2);

            result[0].Name.Should().Be("Alice");
            result[0].Address.Should().NotBeNull();
            result[0].Address.Street.Should().Be("123 Main St");
            result[0].Address.City.Should().Be("London");

            result[1].Name.Should().Be("Bob");
            result[1].Address.Street.Should().Be("456 Elm St");
            result[1].Address.City.Should().Be("Paris");
        }

        [Fact]
        public void ToList_without_nested_columns_should_still_work()
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPlus.Core.Extensions");
            using var package = new ExcelPackage();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

            ws.Cells[1, 1].Value = "Name";
            ws.Cells[1, 2].Value = "Gender";
            ws.Cells[2, 1].Value = "Alice";
            ws.Cells[2, 2].Value = "Female";

            List<NamedMap> result = ws.ToList<NamedMap>();

            result.Should().HaveCount(1);
            result[0].FirstName.Should().Be("Alice");
        }
    }
}
