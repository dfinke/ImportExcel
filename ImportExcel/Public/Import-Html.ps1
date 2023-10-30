
function Import-Html {
    [CmdletBinding()]
    param(
        $Url,
        [int]$Index = 0,
        $Header,
        [int]$FirstDataRow = 0,
        [Switch]$UseDefaultCredentials
    )

    $xlFile = [System.IO.Path]::GetTempFileName() -replace "tmp", "xlsx"
    Remove-Item $xlFile -ErrorAction Ignore

    Write-Verbose "Exporting to Excel file $($xlFile)"

    $data = Get-HtmlTable -Url $Url -TableIndex $Index -Header $Header -FirstDataRow $FirstDataRow -UseDefaultCredentials: $UseDefaultCredentials

    $data | Export-Excel $xlFile -Show -AutoSize
}