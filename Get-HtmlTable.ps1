function Get-HtmlTable {
	param(
		[Parameter(Mandatory = $true)]
		$url,
		$tableIndex = 0,
		$Header,
		[int]$FirstDataRow = 0,
		[Switch]$UseDefaultCredentials
	)
	if ($url -match '^(http:\/\/www\.|https:\/\/www\.|http:\/\/|https:\/\/|www\.)') {
		$r = Invoke-WebRequest $url -UseDefaultCredentials: $UseDefaultCredentials

		$table = $r.ParsedHtml.getElementsByTagName("table")[$tableIndex]
	}
	else {
		$r = New-Object -ComObject "HTMLFile"
		$source = Get-Content -Path $url -Raw
		$r.IHTMLDocument2_write($source)
		
		$table = $r.getElementsByTagName("table").Item($tableIndex)
	}
	
	$propertyNames = $Header

	for ($idx = $FirstDataRow; $idx -lt @($table.rows).count; $idx++) {

		$row = $table.rows.item($idx)
		$cells = @($row.cells)

		if (!$propertyNames) {
			if ($cells[0].tagName -eq 'TD') {
				$propertyNames = @($cells | ForEach-Object {$_.innertext -replace ' ', ''})
			}
			else {
				$propertyNames = @(1..($cells.Count + 2) | ForEach-Object { "P$_" })
			}
			continue
		}

		$result = [ordered]@{}

		for ($counter = 0; $counter -lt $cells.Count; $counter++) {
			$propertyName = $propertyNames[$counter]

			if (!$propertyName) { $propertyName = '[missing]'}
			$result.$propertyName = $cells[$counter].InnerText
		}

		[PSCustomObject]$result
	}
}

