
function Import-Html {
    [CmdletBinding()]
    param(
        $url,        
        $index,
        $Header
    )
    
    $xlFile = [System.IO.Path]::GetTempFileName() -replace "tmp","xlsx"
    rm $xlFile -ErrorAction Ignore

    Write-Verbose "Exporting to Excel file $($xlFile)"

    $data = Get-HtmlTable -url $url -tableIndex $index -Header $Header
    
    $data | Export-Excel $xlFile -Show -AutoSize
}