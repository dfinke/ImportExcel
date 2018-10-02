﻿# To ship, is to choose 

#Import-Module .\ImportExcel.psd1 -Force

$pt=[ordered]@{}

$pt.ServiceInfo=@{    
    SourceWorkSheet='Services'
    PivotRows = "Status"    
    PivotData= @{'Status'='count'}
    IncludePivotChart=$true
    ChartType='BarClustered3D'
}

$pt.ProcessInfo=@{
    SourceWorkSheet='Processes'
    PivotRows = "Company"    
    PivotData= @{'Company'='count'}
    IncludePivotChart=$true
    ChartType='PieExploded3D'
}

$gsv=Get-Service | Select-Object status, Name, displayName, starttype
$ps=Get-Process | Select-Object Name,Company, Handles

$file = "c:\temp\testPT.xlsx"
rm $file -ErrorAction Ignore

$gsv| Export-Excel -Path $file -AutoSize -WorkSheetname Services
$ps | Export-Excel -Path $file -AutoSize -WorkSheetname Processes -PivotTableDefinition $pt -Show 
