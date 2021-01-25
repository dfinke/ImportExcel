function Add-ExcelName {
    [CmdletBinding()]
    param(
        #The range of cells to assign as a name.
        [Parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelRange]$Range,
        #The name to assign to the range. If the name exists it will be updated to the new range. If no name is specified, the first cell in the range will be used as the name.
        [String]$RangeName,
        #targeting a worksheet scope prevents using the Named Range in data validation features.  In Excel the default scope is Workbook
        [switch]$ForceSheetScope
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
        if ($ForceSheetScope) {
            if ($ws.names[$RangeName]) {
                Write-verbose -Message "Updating Named range (worksheet scope) '$RangeName' to $($Range.FullAddressAbsolute)."
                $ws.Names[$RangeName].Address = $Range.FullAddressAbsolute
            }
            else  {
                Write-verbose -Message "Creating Named range (worksheet scope) '$RangeName' as $($Range.FullAddressAbsolute)."
                $null = $ws.Names.Add($RangeName, $Range)
            }
        }
        else {
            $wb = $ws.WorkBook
            if ($wb.names[$RangeName]) {
                Write-verbose -Message "Updating Named range (workbook scope) '$RangeName' to $($Range.FullAddressAbsolute)."
                $wb.Names[$RangeName].Address = $Range.FullAddressAbsolute
            }
            else  {
                Write-verbose -Message "Creating Named range (workbook scope) '$RangeName' as $($Range.FullAddressAbsolute)."
                $null = $wb.Names.Add($RangeName, $Range)
            }
        }
    }
    catch {Write-Warning -Message "Failed adding named range '$RangeName' to worksheet '$($ws.Name)': $_"  }
}
