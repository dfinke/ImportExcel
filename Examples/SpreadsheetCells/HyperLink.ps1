try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

Remove-Item "$env:temp\hyperlink.xlsx" -ErrorAction SilentlyContinue

$(
    New-PSItem '=Hyperlink("https://dfinke.github.io","Doug Finke")' @("Link")
    New-PSItem '=Hyperlink("http://blogs.msdn.com/b/powershell/","PowerShell Blog")'
    New-PSItem '=Hyperlink("http://blogs.technet.com/b/heyscriptingguy/","Hey, Scripting Guy")'

) | Export-Excel "$env:temp\hyperlink.xlsx" -AutoSize -Show
