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
    if ($PSVersionTable.PSVersion.Major -gt 5 -and -not (Get-Command ConvertFrom-Html -ErrorAction SilentlyContinue)) {
         # Invoke-WebRequest on .NET core doesn't have ParsedHtml so we need HtmlAgilityPack or similiar Justin Grote's PowerHTML wraps that nicely
         throw "This version of PowerShell needs the PowerHTML module to process HTML Tables."
    }

    $r = Invoke-WebRequest $Url -UseDefaultCredentials: $UseDefaultCredentials
    $propertyNames = $Header

    if ($PSVersionTable.PSVersion.Major -le 5) {
        $table = $r.ParsedHtml.getElementsByTagName("table")[$TableIndex]
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
    else {
        $h    = ConvertFrom-Html -Content $r.Content
        if ($TableIndex -is [valuetype]) { $TableIndex += 1}
        $rows =    $h.SelectNodes("//table[$TableIndex]//tr")
        if (-not $rows) {Write-Warning "Could not find rows for `"//table[$TableIndex]`" in $Url ."}
        if ( -not  $propertyNames) {
            if (   $tableHeaders  = $rows[$FirstDataRow].SelectNodes("th")) {
                   $propertyNames = $tableHeaders.foreach({[System.Web.HttpUtility]::HtmlDecode( $_.innerText ) -replace '\W+','_' -replace '(\w)_+$','$1' })
                   $FirstDataRow += 1
            }
            else {
                   $c = 0
                   $propertyNames = $rows[$FirstDataRow].SelectNodes("td") | Foreach-Object { "P$c" ; $c ++ }
            }
        }
        Write-Verbose ("Property names: " + ($propertyNames -join ","))
        foreach ($n in $FirstDataRow..($rows.Count-1)) {
            $r      = $rows[$n].SelectNodes("td|th")
            if ($r -and $r.innerText -ne "" -and $r.count -gt $rows[$n].SelectNodes("th").count  ) {
                $c      = 0
                $newObj = [ordered]@{}
                foreach ($p in $propertyNames) {
                    $n  = $null
                    #Join descentandts for cases where the text in the cell is split (e.g with a <BR> ). We also want to remove HTML codes, trim and convert unicode minus sign to "-"
                    $cellText = $r[$c].Descendants().where({$_.NodeType -eq "Text"}).foreach({[System.Web.HttpUtility]::HtmlDecode( $_.innerText ).Trim()}) -Join " " -replace "\u2212","-"
                    if ([double]::TryParse($cellText, [ref]$n)) {$newObj[$p] = $n     }
                    else                                        {$newObj[$p] = $cellText }
                    $c ++
                }
                [pscustomObject]$newObj
            }
        }
    }
}
