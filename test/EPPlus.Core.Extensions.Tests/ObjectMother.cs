using System;
using System.ComponentModel.DataAnnotations;
using System.Drawing;

using EPPlus.Core.Extensions.Configuration;

namespace EPPlus.Core.Extensions.Tests
{
    internal enum Manufacturers
    {
        Opel = 1,
        Ford,
        Mercedes
    }

    internal class WrongCars : IExcelExportable
    {
        [ExcelTableColumn("License plate")]
        public string LicensePlate { get; set; }

        [ExcelTableColumn]
        public Manufacturers Manufacturer { get; set; }

        [ExcelTableColumn("Manufacturing date")]
        public DateTime ManufacturingDate { get; set; }

        [ExcelTableColumn("Is ready for traffic?")]
        public bool Ready { get; set; }
    }

    internal class DefaultMap
    {
        [ExcelTableColumn]
        public string Name { get; set; }

        [ExcelTableColumn]
        public string Gender { get; set; }
    }

    internal class NamedMap
    {
        [ExcelTableColumn("Name")]
        public string FirstName { get; set; }

        [ExcelTableColumn("Gender")]
        public string Sex { get; set; }

        public string NotMapped { get; set; }
    }

    internal class IndexMap
    {
        [ExcelTableColumn(1)]
        public string Name { get; set; }

        [ExcelTableColumn(3)]
        public string Gender { get; set; }
    }

    internal enum Genders
    {
        MALE = 1,
        FEMALE = 2
    }

    internal class EnumStringMap
    {
        [ExcelTableColumn("Name")]
        public string Name { get; set; }

        [ExcelTableColumn("Gender")]
        public Genders Gender { get; set; }
    }

    internal enum Classes : byte
    {
        Ten = 10,
        Nine = 9
    }

    internal class EnumByteMap
    {
        [ExcelTableColumn]
        public string Name { get; set; }

        [ExcelTableColumn]
        public Classes Class { get; set; }
    }

    internal class MultiMap
    {
        [ExcelTableColumn]
        public string Name { get; set; }

        [ExcelTableColumn("Class")]
        public Classes Class { get; set; }

        [ExcelTableColumn("Class")]
        public int ClassAsInt { get; set; }
    }

    internal class DateMap
    {
        [ExcelTableColumn]
        [MinLength(1)]
        [MaxLength(50)]
        public string Name { get; set; }

        [ExcelTableColumn]
        public Genders Gender { get; set; }

        [ExcelTableColumn("Birth date")]
        public DateTime BirthDate { get; set; }

        public int NotMappedProperty { get; set; }
    }

    internal class EnumFailMap
    {
        [ExcelTableColumn]
        public string Name { get; set; }

        [ExcelTableColumn("Gender")]
        public Classes Gender { get; set; }
    }

    internal class CarNullable
    {
        [ExcelTableColumn("Car name")]
        public string Name { get; set; }

        [ExcelTableColumn]
        public int? Price { get; set; }
    }

    internal class StocksNullable
    {
        [ExcelTableColumn(1)]
        public string Barcode { get; set; }

        [ExcelTableColumn(2)]
        public int? Quantity { get; set; }

        [ExcelTableColumn(3)]
        public DateTime? UpdatedDate { get; set; }
    }

    internal class StocksValidation
    {
        [ExcelTableColumn(1)]
        [MinLength(1)]
        [MaxLength(255)]
        public string Barcode { get; set; }

        [ExcelTableColumn(2)]
        [Range(10, int.MaxValue, ErrorMessage = "Please enter a value bigger than {1}")]
        public int Quantity { get; set; }

        [ExcelTableColumn(3)]
        public DateTime UpdatedDate { get; set; }

        public string NotMappedProperty { get; set; }
    }

    internal enum Manufacturers2
    {
        Opel = 1,
        Ford,
        Toyota
    }

    internal class Cars
    {
        [ExcelTableColumn(1)]
        public string LicensePlate { get; set; }

        [ExcelTableColumn]
        public Manufacturers2 Manufacturer { get; set; }

        [ExcelTableColumn("Manufacturing date")]
        public DateTime? ManufacturingDate { get; set; }

        [ExcelTableColumn]
        public int Price { get; set; }

        [ExcelTableColumn]
        public Color Color { get; set; }

        [ExcelTableColumn("Is ready for traffic?")]
        public bool Ready { get; set; }

        public string UnmappedProperty { get; set; }

        public override string ToString()
        {
            return $"{Color} {Manufacturer.ToString()} {ManufacturingDate?.ToShortDateString()}";
        }
    }

    public class Car
    {
        public string Name { get; set; }

        public decimal Price { get; set; }
    }

    public class Person
    {
        public string FirstName { get; set; }

        [ExcelTableColumn(2)]
        public string LastName { get; set; }

        [ExcelTableColumn("Year of Birth")]
        public int YearBorn { get; set; }
    }

    public class ObjectWithoutExcelTableAttributes
    {
        public string FirstName { get; set; }   
       
        public string LastName { get; set; }   
        
        public int YearBorn { get; set; }
    }

    public class ObjectWithWrongAttributeMappings
    {
        [ExcelTableColumn(5)]
        public string FirstName { get; set; }

        [ExcelTableColumn("Firstname")]
        public string LastName { get; set; }

        public int YearBorn { get; set; }
    }
}
