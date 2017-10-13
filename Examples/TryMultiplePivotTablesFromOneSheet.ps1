$file="C:\Temp\test.xlsx"

rm $file -ErrorAction Ignore -Force


$ptd = @{}
$base=@{
    SourceWorkSheet='gsv'
    PivotData= @{'Status'='count'}
    IncludePivotChart=$true
    ChartType='BarClustered3D'
}

$ptd.gpt1 = $base + @{ PivotRows = "ServiceType" }
$ptd.gpt2 = $base + @{ PivotRows = "Status" }
$ptd.gpt3 = $base + @{ PivotRows = "StartType" }
$ptd.gpt4 = $base + @{ PivotRows = "CanStop" }

gsv | Export-Excel -path $file -WorkSheetname gsv -Show -PivotTableDefinition $ptd
