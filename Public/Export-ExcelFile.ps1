function Export-ExcelFile {
    <#
    .SYNOPSIS
        Export dataTable and save it as an Excel xlsx file.

    .DESCRIPTION
        This function applies C# library `ClosedXML`.
        The dataTable can be piped into this function.

    .PARAMETER output_excel_name
        The name of the resulting xlsx file e.g., abc.xlsx.

    .PARAMETER output_sheet_name
        Optional. The default is 'Sheet1'.
        The name of the sheet of the resulting xlsx file.

    .PARAMETER input_excel_name
        Optional. The default is empty.
        Use this if you want the dataTable to be saved in an exisiting xlsx file located in the current directory.

    .PARAMETER data_table
        The input to be saved in xlsx file. Must be dataTable.

    .INPUTS
        data_table.
        input_excel_name (Optional).

    .OUTPUTS
        The resulting xlsx file will be saved in the current directory.

    .EXAMPLE

    #>
    [CmdletBinding()]
    param (
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$output_excel_name,
        [string]$output_sheet_name = 'Sheet1',
        [string]$input_excel_name,
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [System.Data.DataTable]$data_table
    )

    begin {
    }
    process {
        $wb = if ($input_excel_name) { New-Object -TypeName ClosedXML.Excel.XLWorkbook -ArgumentList $(Join-Path $(Get-Location) $input_excel_name) }
              else { New-Object -TypeName ClosedXML.Excel.XLWorkbook }
        $ws = $wb.Worksheets.Add($data_table, $output_sheet_name)
        $ws.Tables.Remove('Table1')
        $wb.SaveAs($(Join-Path $(Get-Location) $output_excel_name))
    }
}