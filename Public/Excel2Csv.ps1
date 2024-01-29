function Excel2Csv {
    <#
    .SYNOPSIS
        Convert Excel (.xlsx, .xlsb, xls) file into csv file.

    .DESCRIPTION
        It can be used without Excel installtion.

    .PARAMETER excel_path
        Path of the Excel file.

    .PARAMETER sheet_names
        The sheets in Excel file you want to convert.
        If more than one, store them in array e.g., @('Sheet1', 'Sheet2').
        These sheets will be saved in separate csv files.
        You can also use wildcard * to indicate you want to include all sheets.

    .INPUTS
        excel_path.

    .OUTPUTS
        csv file(s) saved in the current directory, not necessarily the same directory of Excel file.

    .EXAMPLE

    #>
    [CmdletBinding()]
    param (
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$excel_path,
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string[]]$sheet_names
    )

    $excel_path = (Resolve-Path $excel_path -ErrorAction Stop).Path
    if ([System.IO.Path]::GetExtension($excel_path) -notmatch '^\.xlsx$|^\.xlsb$|^\.xls$') { throw "$excel_path is not a xlsx, xlsb or xls file" }

    if ($sheet_names -eq '*') { $sheet_names = Get-ExcelSheetNames $excel_path }
    $data_dict = Import-ExcelFile -excel_path $excel_path -sheet_names $sheet_names

    foreach ($sheet_name in $sheet_names) {
        $data_dict[$sheet_name] | Export-Csv "$([System.IO.Path]::GetFileNameWithoutExtension($excel_path))`_$sheet_name`.csv" -ErrorAction Stop -Force
    }
}