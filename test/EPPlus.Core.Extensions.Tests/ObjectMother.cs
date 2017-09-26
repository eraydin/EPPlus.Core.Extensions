using System;
using System.Drawing;

namespace EPPlus.Core.Extensions.Tests
{
    enum Manufacturers { Opel = 1, Ford, Mercedes };
    class WrongCars
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

    class DefaultMap
    {
        [ExcelTableColumn]
        public string Name { get; set; }

        [ExcelTableColumn]
        public string Gender { get; set; }
    }

    class NamedMap
    {
        [ExcelTableColumn("Name")]
        public string FirstName { get; set; }

        [ExcelTableColumn("Gender")]
        public string Sex { get; set; }
    }

    class IndexMap
    {
        [ExcelTableColumn(1)]
        public string Name { get; set; }

        [ExcelTableColumn(3)]
        public string Gender { get; set; }
    }

    enum Genders { MALE = 1, FEMALE = 2 }
    class EnumStringMap
    {
        [ExcelTableColumn("Name")]
        public string Name { get; set; }

        [ExcelTableColumn("Gender")]
        public Genders Gender { get; set; }
    }

    enum Classes : byte { Ten = 10, Nine = 9 }
    class EnumByteMap
    {
        [ExcelTableColumn]
        public string Name { get; set; }

        [ExcelTableColumn]
        public Classes Class { get; set; }
    }

    class MultiMap
    {
        [ExcelTableColumn]
        public string Name { get; set; }

        [ExcelTableColumn("Class")]
        public Classes Class { get; set; }

        [ExcelTableColumn("Class")]
        public int ClassAsInt { get; set; }
    }

    class DateMap
    {
        [ExcelTableColumn]
        public string Name { get; set; }

        [ExcelTableColumn]
        public Genders Gender { get; set; }

        [ExcelTableColumn("Birth date")]
        public DateTime BirthDate { get; set; }
    }

    class EnumFailMap
    {
        [ExcelTableColumn]
        public string Name { get; set; }

        [ExcelTableColumn("Gender")]
        public Classes Gender { get; set; }
    }

    class CarNullable
    {
        [ExcelTableColumn("Car name")]
        public string Name { get; set; }

        [ExcelTableColumn]
        public int? Price { get; set; }
    }

    class StocksNullable
    {
        [ExcelTableColumn(1)]
        public string Barcode { get; set; }

        [ExcelTableColumn(2)]
        public int? Quantity { get; set; }

        [ExcelTableColumn(3)]
        public DateTime? UpdatedDate { get; set; }
    }

    enum Manufacturers2 { Opel = 1, Ford, Toyota };
    class Cars
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
            return $"{(Color)} {(Manufacturer.ToString())} {(ManufacturingDate?.ToShortDateString())}";
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
        
        public string LastName { get; set; }

        [ExcelTableColumn("Year of Birth")]
        public int YearBorn { get; set; }
    }
}
