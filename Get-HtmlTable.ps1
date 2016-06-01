function Get-HtmlTable {
    param(
        [Parameter(Mandatory=$true)]
        $url,
        $tableIndex=0,
        $Header,
        [int]$FirstDataRow=0,
        [Switch]$UseDefaultCredentials
    )

    $r = Invoke-WebRequest $url -UseDefaultCredentials: $UseDefaultCredentials
    
    $table = $r.ParsedHtml.getElementsByTagName("table")[$tableIndex]
    $propertyNames=$Header
    $totalRows=@($table.rows).count

    for ($idx = $FirstDataRow; $idx -lt $totalRows; $idx++) {

        $row = $table.rows[$idx]
        $cells = @($row.cells)        

        if(!$propertyNames) {
            if($cells[0].tagName -eq 'th') {
                $propertyNames = @($cells | foreach {$_.innertext -replace ' ',''})
            } else  {
                $propertyNames =  @(1..($cells.Count + 2) | % { "P$_" })
            }
            continue
        }        

        $result = [ordered]@{}

        for($counter = 0; $counter -lt $cells.Count; $counter++) {
            $propertyName = $propertyNames[$counter]

            if(!$propertyName) { $propertyName= '[missing]'}
            $result.$propertyName= $cells[$counter].InnerText
        }

        [PSCustomObject]$result
    }
}