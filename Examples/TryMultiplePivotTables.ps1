ipmo .\ImportExcel.psd1 -Force

$file = "c:\temp\testPT.xlsx"
rm $file -ErrorAction Ignore

$pt=[ordered]@{}

$pt.PT1=@{
    PivotRows = "Status"    
    PivotData= @{'Status'='count'}
    IncludePivotChart=$true
}

$pt.PT2=@{
    PivotRows = "StartType"    
    PivotData= @{'StartType'='count'}
    IncludePivotChart=$true
}


$data = gsv | select status, Name, displayName, starttype
$data | Export-Excel -Path $file -Show -PivotTable $pt -AutoSize 