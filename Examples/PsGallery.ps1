
$top1000 = foreach ($p in 1..50) {
    $c =  Invoke-WebRequest -Uri "https://www.powershellgallery.com/packages" -method Post  -Body "q=&sortOrder=package-download-count&page=$p"
    [regex]::Matches($c.Content,'<table class="width-hundred-percent">.*?</table>', [System.Text.RegularExpressions.RegexOptions]::Singleline) | foreach {
        $name = [regex]::Match($_, "(?<=<h1><a href=.*?>).*(?=</a></h1>)").value
        $n =    [regex]::replace($_,'^.*By:\s*<li role="menuitem">','', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $n =    [regex]::replace($n,'</div>.*$','', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        $by =   [regex]::match($n,'(?<=">).*(?=</a>)').value
        $qty =  [regex]::match($n,'\S*(?= downloads)').value
        [PSCustomObject]@{
            Name = $name
            by   = $by
            Downloads = $qty
        }
    }
}

del "~\Documents\gallery.xlsx"
$pivotdef = New-PivotTableDefinition -PivotTableName 'Summary' -PivotRows by -PivotData @{name="Count"
                                     Downloads="Sum"} -PivotDataToColumn -Activate -ChartType ColumnClustered -PivotNumberFormat '#,###'
$top1000 | export-excel -path '~\Documents\gallery.xlsx' -Numberformat '#,###' -PivotTableDefinition $pivotdef -TableName 'TopDownloads' -Show