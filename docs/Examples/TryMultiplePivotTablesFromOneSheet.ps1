try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$file = "C:\Temp\test.xlsx"

Remove-Item $file -ErrorAction Ignore -Force

$base = @{
    SourceWorkSheet   = 'gsv'
    PivotData         = @{'Status' = 'count'}
    IncludePivotChart = $true
    # ChartType         = 'BarClustered3D'
}

$ptd = [ordered]@{}

# $ptd.gpt1 = $base + @{ PivotRows = "ServiceType" }
# $ptd.gpt2 = $base + @{ PivotRows = "Status" }
# $ptd.gpt3 = $base + @{ PivotRows = "StartType" }
# $ptd.gpt4 = $base + @{ PivotRows = "CanStop" }

$ptd += New-PivotTableDefinition @base servicetype -PivotRows servicetype -ChartType Area3D
$ptd += New-PivotTableDefinition @base status -PivotRows status -ChartType PieExploded3D
$ptd += New-PivotTableDefinition @base starttype -PivotRows starttype -ChartType BarClustered3D
$ptd += New-PivotTableDefinition @base canstop -PivotRows canstop -ChartType ConeColStacked

Get-Service | Export-Excel -path $file -WorkSheetname gsv -Show -PivotTableDefinition $ptd