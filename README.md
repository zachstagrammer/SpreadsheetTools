# SpreadsheetTools

SpreadsheetTools is a lightweight .NET library that reads .xls and .xlsx Excel files (compatible with all BIFF formats) and maps rows to strongly typed C# class instances using header-based reflection.

Supports .NET Framework 4.8 and .NET 6+.

## Features:
- Supports .xls (BIFF2–BIFF8) and .xlsx
- Maps Excel columns to class properties by name (or its DisplayNameAttribute)
- Skips blank rows
- Configurable header match value (e.g. "ID")

## How to use
```c#
using ExcelToolkit;

var importer = new ExcelImporter();
var filePath = "PersonData.xlsx";
var people = importer.Import<Person>(filePath, "ID");

foreach (var person in people)
{
    Console.WriteLine($"{person.Id}\t{person.FirstName}\t{person.LastName}\t{person.Phone}");
}
```

## Example excel file
| ID   | First Name | Last Name | Phone         |
|------|------------|-----------|---------------|
| 1001 | John       | Doe       | (555) 123-4567|
| 1002 | Jane       | Smith     | (555) 234-5678|
| 1003 | Alice      | Johnson   | (555) 345-6789|
| 1004 | Billy Bob  | Brown     | (555) 456-7890|

## Example class
```c#
public class Person
{
    public string ID { get; set; }

    [DisplayName("First Name")]
    public string FirstName { get; set; }

    [DisplayName("Last Name")]
    public string LastName { get; set; }

    public string Phone { get; set; }
}
```

## Import Methods
ExcelToolkit has 3 Import methods for different file header configurations.
1. Header is first row
```c#
var people = importer.Import<Person>(filePath);
```

2. Header row index is specified
```c#
// Zero-based indexing. 2 = header on 3rd row
var people = importer.Import<Person>(filePath, 2);
```

3. Header column A name is specified
```c#
var people = importer.Import<Person>(filePath, "ID");
```
