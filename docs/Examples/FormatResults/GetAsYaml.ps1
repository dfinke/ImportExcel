param(
    [Alias('FullName')]
    [String[]]$Path
)

if ($PSVersionTable.PSVersion.Major -gt 5 -and -not (Get-Command Format-YAML -ErrorAction SilentlyContinue)) {
    throw "This requires EZOut. Install-Module EZOut -AllowClobber -Scope CurrentUser"
}

Import-Excel $Path | Format-YAML