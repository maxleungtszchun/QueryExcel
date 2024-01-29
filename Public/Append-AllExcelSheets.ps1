function Append-AllExcelSheets {
    <#
    .SYNOPSIS
        Append sheets in a Excel (.xlsx, .xlsb, .xls) file.

    .DESCRIPTION
        Append sheets in a Excel (.xlsx, .xlsb, .xls) file.

    .PARAMETER excel_path
        The path of the Excel file.

    .INPUTS
        excel_path.

    .OUTPUTS
        An appended dataTable.

    .EXAMPLE

    #>
    [CmdletBinding()]
    param (
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$excel_path
    )

    $excel_path = (Resolve-Path $excel_path -ErrorAction Stop).Path
    if ([System.IO.Path]::GetExtension($excel_path) -notmatch '^\.xlsx$|^\.xlsb$|^\.xls$') { throw "$excel_path is not a xlsx, xlsb or xls file" }

    $sheet_names = Get-ExcelSheetNames $excel_path
    $data_dict = Import-ExcelFile -excel_path $excel_path -sheet_names $sheet_names

    $i = 1
    foreach ($sheet_name in $sheet_names) {
        if ($i -eq 1) {
            $data = $data_dict[$sheet_name]
            $i++
        } else {
            $data = (Append-ExcelDataTable $data $data_dict[$sheet_name])['Appended']
        }
    }
    return $data
}