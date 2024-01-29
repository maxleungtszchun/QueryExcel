function xls2xlsx {
    <#
    .SYNOPSIS
        Convert xls file into xlsx file.

    .DESCRIPTION
        This function can be run in non-Windows OS because COM is not required. Excel installation is also not required.
        In contrast, `ConvertTo-ExcelXlsx` from PowerShell library `ImportExcel` uses COM and thus can only be run on Windows OS.
        Improving this is one of my motivation of writing this function.

    .PARAMETER xls_path
        Path of the xls file.

    .PARAMETER only_visible
        Optional switch.
        -only_visible means only visible sheets in the xls file will be saved in the resulting xlsx file.

    .INPUTS
        xls_path.

    .OUTPUTS
        xlsx file saved in the current directory, not necessarily the same directory of xls file.

    .EXAMPLE

    #>
    [CmdletBinding()]
    param (
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$xls_path,
        [switch]$only_visible
    )

    $xls_path = (Resolve-Path $xls_path -ErrorAction Stop).Path
    if ([System.IO.Path]::GetExtension($xls_path) -ne '.xls') { throw "$xls_path is not a xls file" }

    $data_dict = if ($only_visible) { Import-ExcelFile $xls_path * -only_visible }
                 else { Import-ExcelFile $xls_path * }

    $file_name = [System.IO.Path]::GetFileNameWithoutExtension($xls_path)

    $i = 1
    foreach ($sheet_name in $data_dict.Keys) {
        if ($i -eq 1) {
            , $data_dict[$sheet_name] | Export-ExcelFile -output_excel_name "$file_name`.xlsx" -output_sheet_name $sheet_name
            $i++
        } else {
            , $data_dict[$sheet_name] | Export-ExcelFile -output_excel_name "$file_name`.xlsx" -output_sheet_name $sheet_name -input_excel_name "$file_name`.xlsx"
        }
    }
}