function Invoke-ExcelReZipFile {
    <#
    #>
    param(
        [Parameter(Mandatory)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage
    )   

    Write-Verbose -Message "Re-Zipping $($ExcelPackage.file) using .NET ZIP library"
    try {
        Add-Type -AssemblyName 'System.IO.Compression.Filesystem' -ErrorAction stop
    }
    catch {
        Write-Error "The -ReZip parameter requires .NET Framework 4.5 or later to be installed. Recommend to install Powershell v4+"
        continue
    }
    try {
        $TempZipPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
        $null = [io.compression.zipfile]::ExtractToDirectory($ExcelPackage.File, $TempZipPath)
        Remove-Item $ExcelPackage.File -Force
        $null = [io.compression.zipfile]::CreateFromDirectory($TempZipPath, $ExcelPackage.File)
        Remove-Item $TempZipPath -Recurse -Force 
    }
    catch { throw "Error resizipping $path : $_" }
}