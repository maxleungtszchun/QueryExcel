function Csv2Excel {
    <#
    .SYNOPSIS
        Convert csv file into xlsx file.

    .DESCRIPTION
        Convert csv file into xlsx file.

    .PARAMETER csv_path
        Path of the csv file.

    .INPUTS
        csv_path.

    .OUTPUTS
        xlsx file saved in the current directory, not necessarily the same directory of csv file.

    .EXAMPLE

    #>
    [CmdletBinding()]
    param (
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$csv_path
    )

    $csv_path = (Resolve-Path $csv_path -ErrorAction Stop).Path
    if ([System.IO.Path]::GetExtension($csv_path) -ne '.csv') { throw "$csv_path is not a csv file" }

    $data_table = Import-Csv -Path $csv_path -Delimiter "," -ErrorAction Stop | Out-DataTable
    , $data_table | Export-ExcelFile -output_excel_name "$([System.IO.Path]::GetFileNameWithoutExtension($csv_path))`.xlsx"
}