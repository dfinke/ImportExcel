
function Import-Html {
    [CmdletBinding()]
    param(
        $url,        
        $index,
        $Header,
        [int]$FirstDataRow=0,
        [Switch]$UseDefaultCredentials,
        [ValidateSet("ById", "ByName", "ByIndex")]
        $SelectionMethod="ByIndex",
        $Selector
    )
    
    $xlFile = [System.IO.Path]::GetTempFileName() -replace "tmp","xlsx"
    rm $xlFile -ErrorAction Ignore

    Write-Verbose "Exporting to Excel file $($xlFile)"

    $data = Get-HtmlTable -url $url -tableIndex $index -Header $Header -FirstDataRow $FirstDataRow -UseDefaultCredentials: $UseDefaultCredentials -SelectionMethod $SelectionMethod -Selector $Selector
    
    $data | Export-Excel $xlFile -Show -AutoSize
}