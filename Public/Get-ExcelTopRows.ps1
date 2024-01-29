function Get-ExcelTopRows {
    <#
    .SYNOPSIS
        Print top 10 rows (default) of a sheet in an Excel (.xlsx or .xlsb or .xls) file.

    .DESCRIPTION
        Print top 10 rows (default) of a sheet in an Excel (.xlsx or .xlsb or .xls) file.

    .PARAMETER excel_path
        Path of the Excel (.xlsx or .xlsb or .xls) file.

    .PARAMETER top_rows
        Optional. Default is 10 i.e., first 10 rows.

    .INPUTS
        excel_path.

    .OUTPUTS
        Top 10 rows of the dataTable.

    .EXAMPLE

    #>
    [CmdletBinding()]
    param (
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$excel_path,
        [ValidateNotNullorEmpty()]
        [string]$sheet_name,
        [ValidateRange(1, 1048576)]
        [int]$top_rows = 10
    )

    $excel_path = (Resolve-Path $excel_path -ErrorAction Stop).Path
    if ([System.IO.Path]::GetExtension($excel_path) -notmatch '^\.xlsx$|^\.xlsb$|^\.xls$') { throw "$excel_path is not a xlsx, xlsb or xls file" }

    if ($sheet_name) {
        $data = (Import-ExcelFile $excel_path $sheet_name)[$sheet_name]
    } else {
        $data_dict = Import-ExcelFile $excel_path
        $data = $data_dict[[string]$data_dict.Keys[0]]
    }

    $total_rows = $data.Rows.Count
    if ($top_rows -gt $total_rows) {
        Write-Warning "Your choice of row ($top_rows) is larger than the total number of record ($total_rows) found in the worksheet. All records are shown"
        return $data
    } else {
        return $data.Rows[0..($top_rows - 1)]
    }
}