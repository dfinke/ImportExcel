function plot {
    param(
        $f,
        $minx,
        $maxx
    )

    $minx=[math]::Round($minx,1)
    $maxx=[math]::Round($maxx,1)
    
    $file = 'C:\temp\plot.xlsx'    
    rm $file -ErrorAction Ignore

    $c = New-ExcelChart -XRange X -YRange Y -ChartType Line -NoLegend -Title Plot -Column 2 -ColumnOffSetPixels 35
    
    $(for ($i = $minx; $i -lt $maxx-.1; $i+=.1) {
        [pscustomobject]@{
            X=$i.ToString("N1")
            Y=(&$f $i)
        }
    }) | Export-Excel $file -Show -AutoNameRange -ExcelChartDefinition $c 
}

function pi {[math]::pi}

plot {[math]::Tan($args[0])} (pi) (3*(pi)/2-.01)