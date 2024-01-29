function Append-AllExcelFiles {
    <#
    .SYNOPSIS
        Append the first sheet of all Excel (.xlsx, .xlsb, .xls) files in a directory.

    .DESCRIPTION
        Append the first sheet of all Excel (.xlsx, .xlsb, .xls) files in a directory.

    .PARAMETER dir_path
        The path of the directory containing at least one Excel file.

    .PARAMETER recurse
        Optional switch.
        -recurse means Excel files in sub-directory will also be considered.

    .INPUTS
        dir_path.

    .OUTPUTS
        An appended dataTable.

    .EXAMPLE

    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$dir_path,
        [switch]$recurse
    )

    $dir_path = (Resolve-Path $dir_path -ErrorAction Stop).Path

    if ($recurse) {
        $excel_files = Get-ChildItem -Path $dir_path -Filter '*.xls*' -File -Recurse
    } else {
        $excel_files = Get-ChildItem -Path $dir_path -Filter '*.xls*' -File
    }

    $excel_files | % {$i = 1} {
        if ($i -eq 1) {
            $data_dict = Import-ExcelFile $_.FullName
            $data = $data_dict[[string]$data_dict.Keys[0]]
            $i++
        } else {
            $append_data_dict = Import-ExcelFile $_.FullName
            $append_data = $append_data_dict[[string]$append_data_dict.Keys[0]]
            $data = (Append-ExcelDataTable $data $append_data)['Appended']
        }
    }
    return $data
}