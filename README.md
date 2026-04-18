# **EPPlus.Core.Extensions** [![CI](https://github.com/eraydin/EPPlus.Core.Extensions/actions/workflows/ci.yml/badge.svg)](https://github.com/eraydin/EPPlus.Core.Extensions/actions/workflows/ci.yml)

### **Installation** [![NuGet version](https://img.shields.io/nuget/v/EPPlus.Core.Extensions.svg)](https://www.nuget.org/packages/EPPlus.Core.Extensions)

```
PM> Install-Package EPPlus.Core.Extensions
```

### **Dependencies**

**.NET 10.0** — *EPPlus >= 8.5.3*

### **License Setup**

EPPlus 8 requires a license declaration before use. For non-commercial/personal projects:

```cs
ExcelPackage.License.SetNonCommercialPersonal("YourName");
```

For commercial use, see [EPPlus licensing](https://epplussoftware.com/developers/licenseexception).

### **Features**

- Converts `IEnumerable<T>` into an Excel worksheet or package
- Reads data from Excel packages and converts them into a `List<T>`
- Supports data annotations for validation (`[Required]`, `[MaxLength]`, `[Range]`, etc.)
- Fluent API for building multi-worksheet workbooks
- Generates Excel templates from classes marked with `[ExcelWorksheet]`

### **Examples**

#### Mapping a class to Excel columns

Use `[ExcelTableColumn]` to map properties to Excel columns by name or index:

```cs
public class PersonDto
{
    [ExcelTableColumn("First name")]
    [Required(ErrorMessage = "First name cannot be empty.")]
    [MaxLength(50, ErrorMessage = "First name cannot be more than {1} characters.")]
    public string FirstName { get; set; }

    [ExcelTableColumn(columnName = "Last name", isOptional = true)]
    public string LastName { get; set; }

    [ExcelTableColumn(3)]
    [Range(1900, 2050, ErrorMessage = "Please enter a value bigger than {1}")]
    public int YearBorn { get; set; }

    public decimal NotMapped { get; set; }

    [ExcelTableColumn(isOptional = true)]
    public decimal OptionalColumn1 { get; set; }

    [ExcelTableColumn(columnIndex = 999, isOptional = true)]
    public decimal OptionalColumn2 { get; set; }
}
```

#### Reading Excel data into a list

```cs
// From the first worksheet:
List<PersonDto> persons = excelPackage.ToList<PersonDto>(c => c.SkipCastingErrors());

// From a named worksheet:
List<PersonDto> persons = excelPackage.GetWorksheet("Persons").ToList<PersonDto>();
```

#### Writing a list to an Excel package

```cs
// Convert to ExcelPackage
ExcelPackage excelPackage = persons.ToExcelPackage();

// Convert to byte array
byte[] xlsx = persons.ToXlsx();

// Fluent multi-worksheet builder
List<PersonDto> pre50  = persons.Where(x => x.YearBorn < 1950).ToList();
List<PersonDto> post50 = persons.Where(x => x.YearBorn >= 1950).ToList();

ExcelPackage excelPackage = pre50.ToWorksheet("< 1950")
    .WithConfiguration(c => c.WithColumnConfiguration(x => x.AutoFit()))
    .WithColumn(x => x.FirstName, "First Name")
    .WithColumn(x => x.LastName, "Last Name")
    .WithColumn(x => x.YearBorn, "Year of Birth")
    .WithTitle("< 1950")
    .NextWorksheet(post50, "> 1950")
    .WithColumn(x => x.LastName, "Last Name")
    .WithColumn(x => x.YearBorn, "Year of Birth")
    .WithTitle("> 1950")
    .ToExcelPackage();
```

#### Generating an Excel template from a class

Mark a class with `[ExcelWorksheet]` to generate a structured template:

```cs
[ExcelWorksheet("Stocks")]
public class StocksDto
{
    [ExcelTableColumn("SKU")]
    public string Barcode { get; set; }

    [ExcelTableColumn]
    public int Quantity { get; set; }
}

// Generate as ExcelPackage
ExcelPackage excelPackage = Assembly.GetExecutingAssembly().GenerateExcelPackage(nameof(StocksDto));

// Generate as ExcelWorksheet inside an existing package
using var excelPackage = new ExcelPackage();
ExcelWorksheet worksheet = excelPackage.GenerateWorksheet(Assembly.GetExecutingAssembly(), nameof(StocksDto));
```
