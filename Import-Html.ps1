function Import-Html {
    [CmdletBinding()]
    param(
        $url,        
        $index
    )
    
    $xlFile = (New-TemporaryFile).fullname -replace "tmp","xlsx"
    rm $xlFile -ErrorAction Ignore

    Write-Verbose "Exporting to Excel file $($xlFile)"

    Get-HtmlTable $url $index |
        Export-Excel $xlFile -Show -AutoSize
}