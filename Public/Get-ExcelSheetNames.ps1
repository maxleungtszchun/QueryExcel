function Get-ExcelSheetNames {
    <#
    .SYNOPSIS
        Get all sheet names of an Excel (.xlsx, .xlsb, .xls) files.

    .DESCRIPTION
        Get sheet names of an Excel (.xlsx, .xlsb, .xls) files.
        C# library `ExcelDataReader` is applied here.

    .PARAMETER excel_path
        Path of the Excel (.xlsx or .xlsb or .xls) file.

    .PARAMETER only_visible
        Optional switch.
        -only_visible means all visible sheets will be returned

    .INPUTS
        excel_path.

    .OUTPUTS
        System.Array of sheet names.

    .EXAMPLE

    #>
    [CmdletBinding()]
    param (
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$excel_path,
        [switch]$only_visible
    )

    $excel_path = (Resolve-Path $excel_path -ErrorAction Stop).Path
    if ([System.IO.Path]::GetExtension($excel_path) -notmatch '^\.xlsx$|^\.xlsb$|^\.xls$') { throw "$excel_path is not a xlsx, xlsb or xls file" }

    $mode = [System.IO.FileMode]::Open
    $access = [System.IO.FileAccess]::Read

    $file_stream = New-Object -TypeName System.IO.FileStream $excel_path, $mode, $access
    $excel_data_reader = [ExcelDataReader.ExcelReaderFactory]::CreateReader($file_stream)

    function FilterSheetCallback {
        param ($reader, $sheet_index)
        return $reader.VisibleState -eq "visible"
    }

    $data_set_conf = New-Object -TypeName ExcelDataReader.ExcelDataSetConfiguration
    if ($only_visible) { $data_set_conf.FilterSheet = $Function:FilterSheetCallback }
    $excel_data_set = [ExcelDataReader.ExcelDataReaderExtensions]::AsDataSet($excel_data_reader, $data_set_conf)

    $excel_data_reader.Dispose()
    $file_stream.Close()
    $file_stream.Dispose()

    return $excel_data_set.Tables.TableName
}