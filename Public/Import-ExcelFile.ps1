function Import-ExcelFile {
    <#
    .SYNOPSIS
        Import Excel file (.xlsx, .xlsb, .xls) to PowerShell.

    .DESCRIPTION
        This function applies C# library `ExcelDataReader`.
        The data in Excel sheets are assumed to be "normal" e.g., start at first row and first column.
        It runs much faster, especially for files with few hundred thousand rows, than the `Import-Excel` function in
        `ImportExcel` PowerShell library although the latter has much more functionality and flexibility.
        The speed consideration is one of my motivation to write this function.

    .PARAMETER excel_path
        Path of your Excel (.xlsx or .xlsb or .xls) file.

    .PARAMETER sheet_names
        The sheets in Excel file you want to include.
        If more than one, store them in array e.g., @('Sheet1', 'Sheet2').
        You can also use wildcard * to indicate you want to include all sheets

    .PARAMETER only_visible
        Optional switch.
        This switch will only be applied if sheet_names is *. -only_visible means all visible sheets in the Excel will be loaded.

    .PARAMETER return_array
        Optional switch.
        -return_array means data of a particular sheet will be saved in System.Array, instead of DataTable (Default).

    .INPUTS
        excel_path.

    .OUTPUTS
        An ordered dictionary (i.e., ordered hash table). Key = sheet name; Value = sheet data.
        Sheet data can be saved in DataTable (Default) or System.Array.

    .EXAMPLE

    #>
    [CmdletBinding()]
    param (
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$excel_path,
        [ValidateNotNullorEmpty()]
        [string[]]$sheet_names,
        [switch]$only_visible,
        [switch]$return_array
    )

    $excel_path = (Resolve-Path $excel_path -ErrorAction Stop).Path
    if ([System.IO.Path]::GetExtension($excel_path) -notmatch '^\.xlsx$|^\.xlsb$|^\.xls$') { throw "$excel_path is not a xlsx, xlsb or xls file" }

    $mode = [System.IO.FileMode]::Open
    $access = [System.IO.FileAccess]::Read

    $file_stream = New-Object -TypeName System.IO.FileStream $excel_path, $mode, $access
    $excel_data_reader = [ExcelDataReader.ExcelReaderFactory]::CreateReader($file_stream)

    if ($only_visible) {
        $all_sheet_names = Get-ExcelSheetNames $excel_path -only_visible
    } else {
        $all_sheet_names = Get-ExcelSheetNames $excel_path
    }

    if ($sheet_names -eq '*') {
        $sheet_names = $all_sheet_names
    } elseif (-not $sheet_names) {
        $sheet_names = $all_sheet_names[0]
    }

    $sheet_index_array = @()
    foreach ($sheet_name in $sheet_names) {
        if ($all_sheet_names.Contains($sheet_name)) {
            $sheet_index_array += $all_sheet_names.IndexOf($sheet_name)
        } else {
            throw "$sheet_name is not found in the Excel workbook / $sheet_name is invisible but you select ``only visible`` option"
        }
    }

    function FilterSheetCallback {
        param ($reader, $sheet_index)
        return $sheet_index_array.Contains($sheet_index)
    }

    function ConfigureDataTableCallback {
        param ($reader)
        $data_table_conf = New-Object -TypeName ExcelDataReader.ExcelDataTableConfiguration
        $data_table_conf.UseHeaderRow = $true
        return $data_table_conf
    }

    $data_set_conf = New-Object -TypeName ExcelDataReader.ExcelDataSetConfiguration
    $data_set_conf.UseColumnDataType = $true
    $data_set_conf.FilterSheet = $Function:FilterSheetCallback
    $data_set_conf.ConfigureDataTable = $Function:ConfigureDataTableCallback

    $excel_data_set = [ExcelDataReader.ExcelDataReaderExtensions]::AsDataSet($excel_data_reader, $data_set_conf)

    $excel_data_reader.Dispose()
    $file_stream.Close()
    $file_stream.Dispose()

    $data_table_collections = $excel_data_set.Tables
    $selected_sheet_names = $data_table_collections.TableName

    $dict = [ordered]@{}
    foreach ($selected_sheet_name in $selected_sheet_names) {
        $selected_sheet_index = $selected_sheet_names.IndexOf($selected_sheet_name)
        if ($return_array) {
            $dict[$selected_sheet_name] = [System.Array]$data_table_collections[$selected_sheet_index]
        } else {
            $dict[$selected_sheet_name] = $data_table_collections[$selected_sheet_index]
        }
    }

    return $dict
}