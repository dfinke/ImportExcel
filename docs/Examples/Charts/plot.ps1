try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

function plot {
    param(
        $f,
        $minx,
        $maxx
    )

    $minx=[math]::Round($minx,1)
    $maxx=[math]::Round($maxx,1)

    #Get rid of pre-exisiting sheet
    $xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
    Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
    Remove-Item $xlSourcefile -ErrorAction Ignore

   # $c = New-ExcelChart -XRange X -YRange Y -ChartType Line -NoLegend -Title Plot -Column 2 -ColumnOffSetPixels 35

    $(for ($i = $minx; $i -lt $maxx-.1; $i+=.1) {
        [pscustomobject]@{
            X=$i.ToString("N1")
            Y=(&$f $i)
        }
    }) | Export-Excel $xlSourcefile -Show -AutoNameRange -LineChart -NoLegend  #-ExcelChartDefinition $c
}

function pi {[math]::pi}

plot -f {[math]::Tan($args[0])} -minx (pi) -maxx (3*(pi)/2-.01)