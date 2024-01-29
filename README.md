# QueryExcel
Query Excel with Powershell

## Getting Started / Examples

### Export data to Excel file with 2 tabs
```powershell
ConvertFrom-Csv @"
Region,State,Units,Price
West,Texas,927,923.71
North,Tennessee,466,770.67
East,Florida,520,458.68
East,Maine,828,661.24
West,Virginia,465,053.58
North,Missouri,436,235.67
South,Kansas,214,992.47
North,North Dakota,789,640.72
South,Delaware,712,508.55
"@ | Out-DataTable | Export-ExcelFile 'salesData.xlsx' 'first_tab'

ConvertFrom-Csv @"
Region,State,Units,Price
West,Texas,927,923.71
West,Virginia,465,053.58
"@ | Out-DataTable | Export-ExcelFile 'salesData.xlsx' 'second_tab' 'salesData.xlsx'
```

### View top 5 rows of the 2 tabs
```powershell
Get-ExcelTopRows '.\salesData.xlsx' 'first_tab' 5
Get-ExcelTopRows '.\salesData.xlsx' 'second_tab' 5
```

### Query the Excel file with SQLite SQL statement
Note -override_existing_temp_db will remove the temporary SQLite database saved during the previous query of the same Excel file.
Query-Excel depends on another PowerShell Module `PSSQLite`. Visit https://github.com/RamblingCookieMonster/PSSQLite for more details.
```powershell
Query-Excel '.\salesData.xlsx' * "SELECT Region, AVG(Units) AS avg FROM salesData_first_tab GROUP BY Region"

Query-Excel '.\salesData.xlsx' * "SELECT * FROM salesData_second_tab" -override_existing_temp_db

Query-Excel '.\salesData.xlsx' * "
    SELECT a.* FROM salesData_first_tab a WHERE a.State IN
        (
            SELECT b.State FROM salesData_second_tab b WHERE b.Units =
            (
                SELECT MAX(Units) from salesData_second_tab
            )
        )"
```
### Import Excel file to PowerShell
```powershell
$data = Import-ExcelFile '.\salesData.xlsx' 'first_tab'
$data['first_tab']
```

### Append all tabs in an Excel file
```powershell
Append-AllExcelSheets '.\salesData.xlsx'
```

### Convert Excel file into csv file
```powershell
Excel2Csv '.\salesData.xlsx' *
```

### Convert csv file into Excel file
```powershell
Csv2Excel '.\salesData_first_tab.csv'
```






