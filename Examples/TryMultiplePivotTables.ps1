# To ship, is to choose 

ipmo .\ImportExcel.psd1 -Force


$pt=[ordered]@{}

$pt.PT1=@{
    
    SourceWorkSheet='Sheet1'
    PivotRows = "Status"    
    PivotData= @{'Status'='count'}
    IncludePivotChart=$true
    ChartType='BarClustered3D'
}

$pt.PT2=@{
    SourceWorkSheet='Sheet2'
    PivotRows = "Company"    
    PivotData= @{'Company'='count'}
    IncludePivotChart=$true
    ChartType='PieExploded3D'
}

$gsv=Get-Service | Select-Object status, Name, displayName, starttype
$ps=Get-Process | Select-Object Name,Company, Handles

$file = "c:\temp\testPT.xlsx"
rm $file -ErrorAction Ignore

$gsv| Export-Excel -Path $file -AutoSize 
$ps | Export-Excel -Path $file -AutoSize -WorkSheetname Sheet2 -PivotTableDefinition $pt -Show 

return 
Get-Service | 
    select status, Name, displayName, starttype | 
    Export-Excel -Path $file -Show -PivotTable $pt -AutoSize 