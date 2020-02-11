
function Import-Html {
    [CmdletBinding()]
    param(
        $url,
        $index,
        $Header,
        [int]$FirstDataRow=0,
        [Switch]$UseDefaultCredentials
    )

    $xlFile = [System.IO.Path]::GetTempFileName() -replace "tmp","xlsx"
    Remove-Item $xlFile -ErrorAction Ignore

    Write-Verbose "Exporting to Excel file $($xlFile)"

    $data = Get-HtmlTable -url $url -tableIndex $index -Header $Header -FirstDataRow $FirstDataRow -UseDefaultCredentials: $UseDefaultCredentials

    $data | Export-Excel $xlFile -Show -AutoSize
}