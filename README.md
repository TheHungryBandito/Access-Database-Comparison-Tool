# Access-Database-Comparison-Tool
Repository for the database comparison tool for Microsoft Access Engine Databases (.mdb)
This tool compares two .mdb files by:
1. Selecting all user-defined tables from both .mdb files.
2. Selecting all data from the tables.
3. Comparing every row as an object with the Table, Column, and Data
4. Displaying the results using Out-GridView and/or redirecting the results to a .csv file.

Supports both 64-bit and 32-bit architecture.

## Scenario 1
Multiple users begin reporting connection errors within a business critical application.
After initial triage and troubleshooting, you identify a recent change to the .mdb configuration database.
You refer to the technical documentation and discover the .mdb configuration is backed up after every change.

Knowing the location of both the .mdb files you create a copy of both configurations to prevent locking the db while querying.
You run Compare-AccessDB against the copies and identify a value had been incorrectly changed in the "Settings" table representing the server connection URI the application uses.

By comparing the configurations in seconds using this function, you have saved hours of manual database comparsions. 
You report your findings and liase with teams to revert this value, resolving the issue in a timely manner.

## Syntax
```PowerShell
Compare-AccessDB
  [-ReferencePath]  [string] (required)
  [-DifferencePath] [string] (required)
  [-OutputCSVPath]  [string] (optional) <Default = "">
  [-IncludeEqual]   [switch] (optional) <Default = $false>
  [-Configuration]  [string] (optional) <Default = "Default"> {"Default", "32-bit", "64-bit"}   
```
**-ReferencePath**  is the full path to the reference .mdb file.

**-DifferencePath** is the full path to the .mdb file to compare against.

**-OutputCSVPath**  if specified, is the full path of the .csv output (including extension).

**-IncludeEqual**   if specified, output includes values that are equal.

**-Configuration**  if specified, creates a temp .ps1 file to be run in the specified architecture for Microsoft Access.

## Usage Examples

Default Comparison
```PowerShell
$ReferencePath = "C:\Application\config.mdb"
$DifferencePath = "C:\Application\Backup\config.mdb"
Compare-AccessDB -ReferencePath $ReferencePath -DifferencePath $DifferencePath
```

Comparison with both differences and equal values
```PowerShell
$ReferencePath = "C:\Application\config.mdb"
$DifferencePath = "C:\Application\Backup\config.mdb"
Compare-AccessDB -ReferencePath $ReferencePath -DifferencePath $DifferencePath -IncludeEqual
```

64-bit PowerShell with 32-bit Microsoft Access - This will re-run the check from a temporary file using a 32-bit PowerShell Session
```PowerShell
$ReferencePath = "C:\Application\config.mdb"
$DifferencePath = "C:\Application\Backup\config.mdb"
Compare-AccessDB -ReferencePath $ReferencePath -DifferencePath $DifferencePath -Configuration "32-bit"
```

32-bit PowerShell with 64-bit Microsoft Access - This will re-run the check from a temporary file using a 64-bit PowerShell Session
```PowerShell
$ReferencePath = "C:\Application\config.mdb"
$DifferencePath = "C:\Application\Backup\config.mdb"
Compare-AccessDB -ReferencePath $ReferencePath -DifferencePath $DifferencePath -Configuration "64-bit"
```
