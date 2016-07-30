function Copy-ExcelWorkSheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $SourceWorkbook,
        [Parameter(Mandatory=$true)]
        $SourceWorkSheet,
        [Parameter(Mandatory=$true)]
        $DestinationWorkbook,
        $DestinationWorkSheet,
        [Switch]$Show
    )

    Write-Verbose "Copying $($SourceWorkSheet) from $($SourceWorkbook) to $($DestinationWorkSheet) in $($DestinationWorkbook)"

    if(!$DestinationWorkSheet) {
        $DestinationWorkSheet = $SourceWorkSheet
    }

    Import-Excel -Path $SourceWorkbook -WorkSheetname $SourceWorkSheet |
        Export-Excel -Path $DestinationWorkbook -WorkSheetname $DestinationWorkSheet -Show:$Show
}