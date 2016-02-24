function Import-Html {
    [CmdletBinding()]
    param(
        $url,        
        $index
    )
    
    $xlFile = [System.IO.Path]::GetTempFileName() -replace "tmp","xlsx"
    rm $xlFile -ErrorAction Ignore

    Write-Verbose "Exporting to Excel file $($xlFile)"

    Get-HtmlTable $url $index |
        Export-Excel $xlFile -Show -AutoSize
}