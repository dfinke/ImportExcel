$xlFile = ".\testExport.xlsx"

Remove-Item -ErrorAction Ignore $xlFile

$ExportOptions = @{
    Path = $xlFile 
    Show=$true
    IncludePivotTable=$true
    IncludePivotChart=$true
    PivotRows="Company"
    PivotData="PM"
    ChartType="PieExploded3D"
}

Get-Process | 
    Select Company, Name, Handles, PM | 
    Export-Excel @ExportOptions 