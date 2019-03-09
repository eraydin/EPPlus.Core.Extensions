using System;

using EPPlus.Core.Extensions.Attributes;

using FluentAssertions;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelTableColumnAttribute_Tests
    {
        [Fact]
        public void Should_cannot_set_both_column_index_and_name()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            var excelTableColumnAttribute = new ExcelTableColumnAttribute();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action1 = () =>
                             {
                                 excelTableColumnAttribute.ColumnIndex = 100;
                                 excelTableColumnAttribute.ColumnName = "TEST";
                             };

            Action action2 = () =>
                             {
                                 excelTableColumnAttribute.ColumnName = "TEST";
                                 excelTableColumnAttribute.ColumnIndex = 100;
                             };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action1.Should().Throw<InvalidOperationException>().WithMessage($"Cannot set both {nameof(excelTableColumnAttribute.ColumnName)} and {nameof(excelTableColumnAttribute.ColumnIndex)}!");
            action2.Should().Throw<InvalidOperationException>().WithMessage($"Cannot set both {nameof(excelTableColumnAttribute.ColumnName)} and {nameof(excelTableColumnAttribute.ColumnIndex)}!");
        }

        [Fact]
        public void Should_column_index_cannot_be_equal_or_less_than_zero()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            var excelTableColumnAttribute = new ExcelTableColumnAttribute();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action1 = () => { excelTableColumnAttribute.ColumnIndex = 0; };
            Action action2 = () => { excelTableColumnAttribute.ColumnIndex = -10; };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action1.Should().Throw<InvalidOperationException>().WithMessage($"{nameof(excelTableColumnAttribute.ColumnIndex)} cannot be zero or negative!");
            action2.Should().Throw<InvalidOperationException>().WithMessage($"{nameof(excelTableColumnAttribute.ColumnIndex)} cannot be zero or negative!");
        }

        [Fact]
        public void Should_column_name_cannot_be_null_or_empty()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            var excelTableColumnAttribute = new ExcelTableColumnAttribute();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            Action action1 = () => { excelTableColumnAttribute.ColumnName = "   "; };
            Action action2 = () => { excelTableColumnAttribute.ColumnName = null; };

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            action1.Should().Throw<ArgumentException>().WithMessage("Value must not be empty*");
            action2.Should().Throw<ArgumentNullException>().WithMessage("Value cannot be null*");
        }
    }
}
