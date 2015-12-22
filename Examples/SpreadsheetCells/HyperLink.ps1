rm *.xlsx

$(
    New-PSItem '=Hyperlink("http://dougfinke.com/blog","Doug Finke")' @("Link")
    New-PSItem '=Hyperlink("http://blogs.msdn.com/b/powershell/","PowerShell Blog")' @("Link")
    New-PSItem '=Hyperlink("http://blogs.technet.com/b/heyscriptingguy/","Hey, Scripting Guy")' @("Link")
    
) | Export-Excel hyperlink.xlsx -AutoSize -Show 