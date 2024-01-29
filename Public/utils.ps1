function Get-ColumnNames {
    param ($d)
    return $d | Get-Member | Where { $_.MemberType -eq 'Property' } | Select -ExpandProperty Name
}

function Append-ExcelDataTable {
    [CmdletBinding()]
    param(
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [System.Data.DataTable]$a,
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [System.Data.DataTable]$b
    )

    $a_col_names = Get-ColumnNames $a

    $compare_cols = Compare-Object $a_col_names $(Get-ColumnNames $b)
    if ($compare_cols) { Write-Warning $compare_cols; throw 'The Excel files have different columns' }

    foreach ($b_row in $b.Rows) {
        $new_row = $a.NewRow()
        foreach ($col_name in $a_col_names) {
            $new_row[$col_name] = $b_row[$col_name]
        }
        $a.Rows.Add($new_row)
    }
    return @{'Appended' = $a}
}

function Append-ExcelFiles {
    [CmdletBinding()]
    param(
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$excel_path_a,
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$sheet_name_a,
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$excel_path_b,
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$sheet_name_b
    )

    $a = (Import-ExcelFile $excel_path_a $sheet_name_a)[$sheet_name_a]
    $b = (Import-ExcelFile $excel_path_b $sheet_name_b)[$sheet_name_b]

    return (Append-ExcelDataTable $a $b)['Appended']
}