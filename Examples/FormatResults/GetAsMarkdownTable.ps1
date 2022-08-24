param(
    [Alias('FullName')]
    [String[]]$Path
)

if ($PSVersionTable.PSVersion.Major -gt 5 -and -not (Get-Command Format-Markdown -ErrorAction SilentlyContinue)) {
    throw "This requires EZOut. Install-Module EZOut -AllowClobber -Scope CurrentUser"
}

Import-Excel $Path | Format-Markdown 