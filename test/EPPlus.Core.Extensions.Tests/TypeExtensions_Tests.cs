using FluentAssertions;
using System;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class TypeExtensions_Tests
    {
        [Fact]
        public void Test_IsNullable()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Type typeOfNullableInteger = typeof(int?);
            Type typeOfLong = typeof(long);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool typeOfNullableIntegerResult = typeOfNullableInteger.IsNullable();
            bool typeOfLongResult = typeOfLong.IsNullable();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            typeOfNullableIntegerResult.Should().BeTrue();
            typeOfLongResult.Should().BeFalse();
        }

        [Fact]
        public void Test_IsNumeric()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            Type typeOfInteger = typeof(int);
            Type typeOfNullableInteger = typeof(int?);
            Type typeOfString = typeof(string);
            Type typeOfException = typeof(Exception);

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            bool typeOfIntegerResult = typeOfInteger.IsNumeric();
            bool typeOfNullableIntegerResult = typeOfNullableInteger.IsNumeric();
            bool typeOfStringResult = typeOfString.IsNumeric();
            bool typeOfExceptionResult = typeOfException.IsNumeric();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            typeOfIntegerResult.Should().BeTrue();
            typeOfNullableIntegerResult.Should().BeFalse();
            typeOfStringResult.Should().BeFalse();
            typeOfExceptionResult.Should().BeFalse();
        }
    }
}
