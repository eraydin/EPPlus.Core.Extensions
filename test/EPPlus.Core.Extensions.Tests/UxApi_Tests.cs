using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using EPPlus.Core.Extensions.Attributes;
using EPPlus.Core.Extensions.Results;

using FluentAssertions;

using OfficeOpenXml;
using OfficeOpenXml.Table;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class UxApiTests : TestBase
    {
        [Fact]
        public void Try_get_helpers_should_return_found_objects_and_false_for_missing_objects()
        {
            ExcelPackage1.TryGetWorksheet("TEST1", out ExcelWorksheet worksheet).Should().BeTrue();
            worksheet.Should().NotBeNull();

            ExcelPackage1.Workbook.TryGetWorksheet("TEST1", out ExcelWorksheet workbookWorksheet).Should().BeTrue();
            workbookWorksheet.Should().BeSameAs(worksheet);

            ExcelPackage1.TryGetTable("test1", out ExcelTable packageTable).Should().BeTrue();
            packageTable.Should().NotBeNull();

            worksheet.TryGetTable("TEST1", out ExcelTable worksheetTable).Should().BeTrue();
            worksheetTable.Should().BeSameAs(packageTable);

            ExcelPackage1.TryGetWorksheet("missing", out _).Should().BeFalse();
            ExcelPackage1.TryGetTable("missing", out _).Should().BeFalse();
            worksheet.TryGetTable("missing", out _).Should().BeFalse();
        }

        [Fact]
        public void Package_should_read_rows_by_worksheet_name()
        {
            List<DateMap> list = ExcelPackage1.ToListFromWorksheet<DateMap>("TEST1");
            IEnumerable<DateMap> enumerable = ExcelPackage1.AsEnumerableFromWorksheet<DateMap>("TEST1");

            list.Should().HaveCount(5);
            enumerable.Should().HaveCount(5);
        }

        [Fact]
        public void On_row_should_be_a_discoverable_alias_for_intercept()
        {
            List<DateMap> list = ExcelPackage1.ToListFromWorksheet<DateMap>("TEST1", configuration =>
                configuration.OnRow((item, rowIndex) => item.NotMappedProperty = rowIndex));

            list.Should().OnlyContain(item => item.NotMappedProperty > 0);
        }

        [Fact]
        public void Read_should_not_capture_exceptions_thrown_by_user_callbacks()
        {
            ExcelTable table = ExcelPackage1.GetWorksheet("TEST1").GetTable("TEST1");

            Action action = () => table.Read<DateMap>(configuration =>
                configuration.OnRow((item, rowIndex) => throw new InvalidOperationException("Callback failed.")));

            action.Should().Throw<InvalidOperationException>().WithMessage("Callback failed.");
        }

        [Fact]
        public void Read_should_capture_casting_errors_and_keep_partially_mapped_items()
        {
            ExcelTable table = ExcelPackage1.GetWorksheet("TEST1").GetTable("TEST1");

            ExcelReadResult<EnumFailMap> result = table.Read<EnumFailMap>();

            result.Items.Should().HaveCount(5);
            result.Errors.Should().NotBeEmpty();
            result.Errors.Should().OnlyContain(error => error.Kind == ExcelReadErrorKind.Casting);
            result.Errors.Should().OnlyContain(error => error.Context != null && error.Exception != null);
            result.HasErrors.Should().BeTrue();
            result.IsSuccess.Should().BeFalse();
        }

        [Fact]
        public void Read_should_capture_missing_column_mappings()
        {
            ExcelTable table = ExcelPackage1.GetWorksheet("TEST1").GetTable("TEST1");

            ExcelReadResult<ObjectWithWrongAttributeMappings> result = table.Read<ObjectWithWrongAttributeMappings>();

            result.Items.Should().HaveCount(5);
            result.Errors.Should().ContainSingle(error => error.Kind == ExcelReadErrorKind.Mapping);
            result.Errors.Single().Context.PropertyName.Should().Be(nameof(ObjectWithWrongAttributeMappings.LastName));
        }

        [Fact]
        public void Read_should_capture_out_of_range_index_mappings()
        {
            ExcelTable table = ExcelPackage1.GetWorksheet("TEST1").GetTable("TEST1");

            ExcelReadResult<MissingIndexMap> result = table.Read<MissingIndexMap>();

            result.Items.Should().HaveCount(5);
            result.Errors.Should().ContainSingle(error => error.Kind == ExcelReadErrorKind.Mapping);
            result.Errors.Single().Context.PropertyName.Should().Be(nameof(MissingIndexMap.Value));
        }

        [Fact]
        public void Read_should_report_a_mapping_error_when_the_type_has_no_column_attributes()
        {
            ExcelTable table = ExcelPackage1.GetWorksheet("TEST1").GetTable("TEST1");

            ExcelReadResult<ObjectWithoutExcelTableAttributes> result = table.Read<ObjectWithoutExcelTableAttributes>();

            result.Items.Should().BeEmpty();
            result.Errors.Should().ContainSingle(error => error.Kind == ExcelReadErrorKind.Mapping);
        }

        [Fact]
        public void Read_should_capture_data_annotation_validation_errors()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Validation");
            worksheet.Cells[1, 1].Value = "Barcode";
            worksheet.Cells[1, 2].Value = "Quantity";
            worksheet.Cells[1, 3].Value = "UpdatedDate";
            worksheet.Cells[2, 1].Value = "ABC";
            worksheet.Cells[2, 2].Value = 1;
            worksheet.Cells[2, 3].Value = DateTime.Today;

            ExcelReadResult<StocksValidation> result = worksheet.Read<StocksValidation>();

            result.Items.Should().ContainSingle();
            result.Errors.Should().ContainSingle(error => error.Kind == ExcelReadErrorKind.Validation);
            result.Errors.Single().Context.CellAddress.Address.Should().Be("B2");
        }

        [Fact]
        public void Package_read_should_support_named_worksheets_and_header_configuration()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("People");
            worksheet.Cells[1, 1].Value = "Title";
            worksheet.Cells[2, 1].Value = "Name";
            worksheet.Cells[2, 2].Value = "Gender";
            worksheet.Cells[3, 1].Value = "Ada";
            worksheet.Cells[3, 2].Value = "Female";

            ExcelReadResult<NamedMap> result = package.Read<NamedMap>("People", configuration => configuration.WithHeaderRowIndex(2));

            result.IsSuccess.Should().BeTrue();
            result.Items.Should().ContainSingle();
            result.Items[0].FirstName.Should().Be("Ada");
        }

        [Fact]
        public void Export_should_offer_default_and_explicit_worksheet_names()
        {
            var people = new[] { new Person { LastName = "Lovelace", YearBorn = 1815 } };

            using ExcelPackage defaultPackage = people.ToWorksheet().ToExcelPackage();
            defaultPackage.Workbook.Worksheets.Single().Name.Should().Be(nameof(Person));

            byte[] buffer = people.ToXlsx("People", addHeaderRow: true);
            using var namedPackage = new ExcelPackage(new MemoryStream(buffer));
            namedPackage.Workbook.Worksheets.Single().Name.Should().Be("People");
        }

        [Fact]
        public void Export_should_enumerate_source_rows_once()
        {
            var enumerationCount = 0;

            IEnumerable<Person> Rows()
            {
                enumerationCount++;
                yield return new Person { LastName = "Lovelace", YearBorn = 1815 };
                yield return new Person { LastName = "Hopper", YearBorn = 1906 };
            }

            using ExcelPackage package = Rows().ToExcelPackage();

            enumerationCount.Should().Be(1);
            package.Workbook.Worksheets.Single().Dimension.Rows.Should().Be(3);
        }

        [Fact]
        public void Existing_default_literal_calls_should_remain_unambiguous()
        {
            var people = new[] { new Person { LastName = "Lovelace", YearBorn = 1815 } };

            byte[] buffer = people.ToXlsx(default);
            List<DateMap> rows = ExcelPackage1.ToList<DateMap>(default);

            buffer.Should().NotBeEmpty();
            rows.Should().HaveCount(5);
        }

        [Fact]
        public void Epplus_8_6_1_should_calculate_regex_formulas_and_round_trip_workbooks()
        {
            using var package = new ExcelPackage();
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Regex");
            worksheet.Cells["A1"].Formula = "REGEXTEST(\"abc123\",\"[0-9]+\")";

            package.Workbook.Calculate();

            worksheet.Cells["A1"].Value.Should().Be(true);

            byte[] buffer = package.GetAsByteArray();
            using var reopened = new ExcelPackage(new MemoryStream(buffer));
            reopened.Workbook.Worksheets["Regex"].Cells["A1"].Formula.Should().EndWith("REGEXTEST(\"abc123\",\"[0-9]+\")");
        }

        public class MissingIndexMap
        {
            [ExcelTableColumn(999)]
            public string Value { get; set; }
        }
    }
}
