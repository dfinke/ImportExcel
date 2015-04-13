$xlFile = ".\testExport.xlsx"

Remove-Item -ErrorAction Ignore $xlFile

$ExportOptions = @{
    Path              = $xlFile 
    Show              = $true
    IncludePivotTable = $true
    IncludePivotChart = $true
    PivotRows         = echo Company Name    
    PivotData         = "PM"
    ChartType         = "BarClustered3D"
    #Password          = "Test"
}

Get-Process | 
    Where Company |
    Select Company, Name, Handles, PM | 
    Export-Excel @ExportOptions 