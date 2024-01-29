if (-not $PSScriptRoot) { $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path }

foreach ($folder in @('Public')) {
    Get-ChildItem -Path $(Join-Path $PSScriptRoot $folder '*.ps1') | % {
        try { . $_.FullName }
        catch { Write-Error "Fail to import $($_.FullName)" }
    }
}

if ($PSEdition -eq 'Core') {
    if ($isMacOS -and [Environment]::Is64BitOperatingSystem) {
        if ($(bash -c 'uname -m') -eq 'arm64') {
            $pssqlite_path = Join-Path $(Split-Path (Get-Module -ListAvailable PSSQLite).Path) 'core' 'osx-x64' 'SQLite.Interop.dll'
            Copy-Item $(Join-Path $PSScriptRoot 'dll' 'mac_arm64' SQLite.Interop.dll) -Destination $pssqlite_path -Force
            Write-Verbose "``SQLite.Interop.dll`` in ``$pssqlite_path`` is changed to arm version"
        }
    } elseif ($isLinux -and [Environment]::Is64BitOperatingSystem) {
        if ($(bash -c 'uname -m') -eq 'aarch64') {
            $pssqlite_path = Join-Path $(Split-Path (Get-Module -ListAvailable PSSQLite).Path) 'core' 'linux-x64' 'SQLite.Interop.dll'
            Copy-Item $(Join-Path $PSScriptRoot 'dll' 'linux_arm64' SQLite.Interop.dll) -Destination $pssqlite_path -Force
            Write-Verbose "``SQLite.Interop.dll`` in ``$pssqlite_path`` is changed to arm version"
        }
    }
}

# without using Export-ModuleMember() here, all functions (except variables) are exported. FunctionsToExport in QueryExcel.psd1 will the export as well.
