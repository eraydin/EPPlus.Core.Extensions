# **EPPlus.Core.Extensions** [![Build status](https://ci.appveyor.com/api/projects/status/cdhoa8m20k2k71ke?svg=true)](https://ci.appveyor.com/project/eraydin/epplus-core-extensions) [![codecov](https://codecov.io/gh/eraydin/EPPlus.Core.Extensions/graph/badge.svg)](https://codecov.io/gh/eraydin/EPPlus.Core.Extensions)

### **Installation** [![NuGet version](https://badge.fury.io/nu/EPPlus.Core.Extensions.svg)](https://badge.fury.io/nu/EPPlus.Core.Extensions)

It's as easy as `PM> Install-Package EPPlus.Core.Extensions` from [nuget](http://nuget.org/packages/EPPlus.Core.Extensions)

### **Dependencies**

**.NET Framework 4.6.1**
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*EPPlus >= 4.5.3.3* 
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*System.ComponentModel.Annotations >= 4.7.0*

**.NET Standard 2.0**
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*EPPlus >= 4.5.3.3*
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*System.ComponentModel.Annotations >= 4.7.0*

### **Documentation and Examples**

The project will be documented soon but you can look at the test project for now. I hope it has enough number of examples to give you better idea about how to use these extension methods. 

- Converts IEnumerable<T> into an Excel worksheet/package
- Reads data from Excel packages and convert them into a List<T>.

##### Basic examples:

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

        [ExcelTableColumn(columnIndex=999, isOptional = true)]
        public decimal OptionalColumn2 { get; set; }
    }      
```

- Converting from Excel to list of objects

```cs
    // Direct usage: 
        excelPackage.ToList<PersonDto>(configuration => configuration.SkipCastingErrors());

    // Specific worksheet: 
        excelPackage.GetWorksheet("Persons").ToList<PersonDto>(); 
``` 
    
- From a list of objects to Excel package

```cs
    List<PersonDto> persons = new List<PersonDto>();
         
    // Convert list into ExcelPackage
        ExcelPackage excelPackage = persons.ToExcelPackage();

    // Convert list into byte array 
        byte[] excelPackageXlsx = persons.ToXlsx();
       

    // Generate ExcelPackage with configuration

    List<PersonDto> pre50 = persons.Where(x => x.YearBorn < 1950).ToList();
    List<PersonDto> post50 = persons.Where(x => x.YearBorn >= 1950).ToList();
        
    ExcelPackage excelPackage = pre50.ToWorksheet("< 1950")
                             .WithConfiguration(configuration => configuration.WithColumnConfiguration(x => x.AutoFit()))
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

- Generating an Excel template from ExcelWorksheetAttribute marked classes

```cs 
    [ExcelWorksheet("Stocks")]
    public class StocksDto
    {
        [ExcelTableColumn("SKU")]
        public string Barcode { get; set; }
    
        [ExcelTableColumn]
        public int Quantity { get; set; }
    }   

    // To ExcelPackage
    ExcelPackage excelPackage = Assembly.GetExecutingAssembly().GenerateExcelPackage(nameof(StocksDto));
 
    // To ExcelWorksheet
    using(var excelPackage = new ExcelPackage()){ 
    
        ExcelWorksheet worksheet = excelPackage.GenerateWorksheet(Assembly.GetExecutingAssembly(), nameof(StocksDto));
    
    }  
```
