rm *.xlsx

$(
    New-PSItem '=Hyperlink("http://dougfinke.com/blog","Doug Finke")' @("Link")
    New-PSItem '=Hyperlink("http://blogs.msdn.com/b/powershell/","PowerShell Blog")'
    New-PSItem '=Hyperlink("http://blogs.technet.com/b/heyscriptingguy/","Hey, Scripting Guy")'
    
) | Export-Excel hyperlink.xlsx -AutoSize -Show 
