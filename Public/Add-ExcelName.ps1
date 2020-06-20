function Add-ExcelName {
    [CmdletBinding()]
    param(
        #The range of cells to assign as a name.
        [Parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelRange]$Range,
        #The name to assign to the range. If the name exists it will be updated to the new range. If no name is specified, the first cell in the range will be used as the name.
        [String]$RangeName
    )
    try {
        $ws = $Range.Worksheet
        if (-not $RangeName) {
            $RangeName = $ws.Cells[$Range.Start.Address].Value
            $Range  = ($Range.Worksheet.cells[($Range.start.row +1), $Range.start.Column ,  $Range.end.row, $Range.end.column])
        }
        if ($RangeName -match '\W') {
            Write-Warning -Message "Range name '$RangeName' contains illegal characters, they will be replaced with '_'."
            $RangeName = $RangeName -replace '\W','_'
        }
        if ($ws.names[$RangeName]) {
            Write-verbose -Message "Updating Named range '$RangeName' to $($Range.FullAddressAbsolute)."
            $ws.Names[$RangeName].Address = $Range.FullAddressAbsolute
        }
        else  {
            Write-verbose -Message "Creating Named range '$RangeName' as $($Range.FullAddressAbsolute)."
            $null = $ws.Names.Add($RangeName, $Range)
        }
    }
    catch {Write-Warning -Message "Failed adding named range '$RangeName' to worksheet '$($ws.Name)': $_"  }
}
