function Query-Excel {
    <#
    .SYNOPSIS
        Query Excel, csv and tsv file with SQLite SQL statement.

    .DESCRIPTION
        You can query Excel files even without installing MS Excel in your computer.

    .PARAMETER excel_path
        Path of your Excel (.xlsx, .xlsb, .xls) or csv or tsv file.

    .PARAMETER sheet_names
        The sheets in Excel file you want to include.
        If more than one, store them in array e.g., @('Sheet1', 'Sheet2')
        You can also use wildcard * to indicate you want to include all sheets.
        For csv and tsv files, you can specify any string including empty string
        since there is only one sheet in these files and will be selected automatically.

    .PARAMETER sql_statement
        The SQLite SQL statement to query Excel sheets. Visit https://www.sqlite.org/lang.html for more details.
        For Excel file (xlsx, xlsb, xls), the table name in the SQL statement is <file_name_without_extension>_<sheet_name>,
        any space in file name and sheet name should be replaced by underscore.
        e.g., if file name is `abc.xlsx` and sheet name is `Sheet 1`, the table name is `abc_Sheet_1`.

        For csv and tsv file, the table name is simply <file_name_without_extension> without sheet name.
        Spaces should be replaced by underscore.

    .PARAMETER override_existing_temp_db
        Optional switch.
        If you have already run this command on a particular Excel before, its specific SQLite database file is saved
        in the temporary folder. You can use -override_existing_temp_db to override the existing SQLite database.

    .PARAMETER only_visible
        Optional switch.
        This switch will only be effective if sheet_names is *. -only_visible means all visible sheets in the Excel will be loaded.

    .INPUTS
        excel_path.

    .OUTPUTS
        System.Array of the result.

    .EXAMPLE

    #>
    [CmdletBinding()]
    param (
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$excel_path,
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string[]]$sheet_names,
        [ValidateNotNullorEmpty()]
        [Parameter(Mandatory=$true)]
        [string]$sql_statement,
        [switch]$override_existing_temp_db,
        [switch]$only_visible
    )

    function New-TemporaryDirectory {
        param ([string]$hash)
        $parent = [System.IO.Path]::GetTempPath()
        New-Item -ItemType Directory -Path (Join-Path $parent $hash)
    }

    $excel_path = (Resolve-Path $excel_path -ErrorAction Stop).Path
    $file_extension = [System.IO.Path]::GetExtension($excel_path)
    switch -Regex ($file_extension) {
        '^\.csv$' { Write-Warning "$excel_path is a csv file. ``Query-Excel`` supports csv as well"; $sheet_names = '/csv/' }
        '^\.tsv$' { Write-Warning "$excel_path is a tsv file. ``Query-Excel`` supports tsv as well"; $sheet_names = '/tsv/' }
        '^\.xlsx$|^\.xlsb$|^\.xls$' { continue }
        default { throw "$excel_path is not a xlsx, xlsb, xls, csv or tsv file" }
    }

    $stream = [IO.MemoryStream]::new([byte[]][char[]]$excel_path)
    $hashed_value = (Get-FileHash -InputStream $stream -Algorithm SHA256).Hash

    $last_db_dir = Join-Path $([System.IO.Path]::GetTempPath()) $hashed_value
    $last_db_path = Join-Path $last_db_dir tmp_db.SQLite

    if ((Test-Path $last_db_dir) -and (-not $override_existing_temp_db)) {
        if (Test-Path $last_db_path) { $db_path = $last_db_path } else { throw 'temp_db.SQLite is missing in the temporary folder' }
    } else {
        if (Test-Path $last_db_dir) { Remove-Item -Recurse -Force $last_db_dir }
        $tmp_dir = New-TemporaryDirectory $hashed_value
        $db_path = Join-Path $tmp_dir tmp_db.SQLite

        if ($only_visible) {
            if ($sheet_names -eq '*') { $sheet_names = Get-ExcelSheetNames $excel_path -only_visible }
        } else {
            if ($sheet_names -eq '*') { $sheet_names = Get-ExcelSheetNames $excel_path }
        }

        if ($file_extension -match '^\.xlsx$|^\.xlsb$|^\.xls$') {
            try {
                $data_dict = Import-ExcelFile -excel_path $excel_path -sheet_names $sheet_names
            } catch {
                throw "Fail to import Excel file ($($_.Exception.Message))"
            }
        }

        foreach ($sheet_name in $sheet_names) {
            $sql_tbl_name = [System.IO.Path]::GetFileNameWithoutExtension($excel_path).replace(' ', '_')
            switch -Regex ($file_extension) {
                '^\.csv$' {
                    $data = Import-Csv -Path $excel_path -ErrorAction Stop | Out-DataTable
                }
                '^\.tsv$' {
                    $data = Import-Csv -Path $excel_path -Delimiter "`t" -ErrorAction Stop | Out-DataTable
                }
                '^\.xlsx$|^\.xlsb$|^\.xls$' {
                    $data = $data_dict[$sheet_name]
                    $sql_tbl_name = "$sql_tbl_name`_$sheet_name".replace(' ', '_')
                } default {
                    throw 'File type is not supported'
                }
            }

            $col_names = $data | Get-Member | Where { $_.MemberType -eq 'Property' } | Select -ExpandProperty Name
            $sql_col_name = ''
            foreach ($col_name in $col_names) {
                if ($col_name -match '\s') { throw "Column ``$col_name`` cannot include any space" }
                $sql_col_name += " $col_name TEXT,"
            }
            $sql_col_name = $sql_col_name.Substring(1, $sql_col_name.Length - 2)
            $sql = "CREATE TABLE $sql_tbl_name ($sql_col_name)"

            try{
                Invoke-SQLiteQuery -Query $sql -DataSource $db_path
                Invoke-SQLiteBulkCopy -DataTable $data -DataSource $db_path -Table $sql_tbl_name -Force
                Write-Verbose "Imported $sheet_name to temporary SQLite database located at $db_path"
            } catch {
                throw "Fail to create temporary SQLite database ($($_.Exception.Message))"
            }
        }
    }
    try {
        Invoke-SQLiteQuery -DataSource $db_path -Query $sql_statement
    } catch {
        # when error occurs, Invoke-SQLiteQuery simply aborts the shell session. Thus, try-catch doesn't have any effects
        throw $_.Exception.Message
    }
}