# https://www.leeholmes.com/blog/2015/01/05/extracting-tables-from-powershells-invoke-webrequest/
# tweaked from the above code
function Get-HtmlTable {
    param(
        [Parameter(Mandatory=$true)]
        $Url,
        $TableIndex=0,
        $Header,
        [int]$FirstDataRow=0,
        [Switch]$UseDefaultCredentials
    )

    $r = Invoke-WebRequest $Url -UseDefaultCredentials: $UseDefaultCredentials

    $table = $r.ParsedHtml.getElementsByTagName("table")[$TableIndex]
    $propertyNames=$Header
    $totalRows=@($table.rows).count

    for ($idx = $FirstDataRow; $idx -lt $totalRows; $idx++) {

        $row = $table.rows[$idx]
        $cells = @($row.cells)

        if(!$propertyNames) {
            if($cells[0].tagName -eq 'th') {
                $propertyNames = @($cells | ForEach-Object {$_.innertext -replace ' ',''})
            } else  {
                $propertyNames =  @(1..($cells.Count + 2) | Foreach-Object { "P$_" })
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
