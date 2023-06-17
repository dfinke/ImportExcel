function Get-ExcelTable {
	[CmdLetBinding()]
	[OutputType([PSCustomObject])]
	param (
		[Alias('FullName')][string]$Path,
		[OfficeOpenXml.ExcelPackage]$ExcelPackage,
		[Parameter(ValueFromPipeline)]
		$TableName, # input filter by name or index
		[string[]]$WorksheetName, # input filter
		[Alias('data')][switch]$Content, # forces to retrieve table data; by default output is table names
		[Alias('ie')][switch]$IncludeEmptySheet, # input filter
		[Alias('eh')][switch]$ExcludeHiddenSheet, # input filter
		[string]$Password,
		[Alias('RawTable')][switch]$GridTable # experimental; TODO
	)

	begin {
		$tabid = [System.Collections.Generic.List[object]]::new()
	}

	process {$tabid.AddRange(@($TableName))}

	end {
		function ConvertFrom-ExcelColumnName ([string]$columnName) {
			$sum = 0
			$columnName.ToCharArray().ForEach{
				$sum *= 26
				$sum += [char]$_.tostring().toupper() - [char]'A'+1
			}
			$sum
		} # END

		# auto ParameterSet resolver
		$Excel = if ($ExcelPackage) {$ExcelPackage} 
		elseif ($Path) {
			$fPath = (Resolve-Path $Path -ErrorAction SilentlyContinue).ProviderPath
			if (-not $fPath) {
				Write-Warning "'$Path' file not found"
				exit
			}
			$Stream = [System.IO.FileStream]::new($fPath, 'Open', 'Read', 'ReadWrite')
			$pkg = [OfficeOpenXml.ExcelPackage]::new()
			if ($Password) {$pkg.Load($stream, $Password)}
			else {$pkg.Load($stream)}
			$pkg
		} else {return}

		$Worksheets = if ($WorksheetName -and $WorksheetName -ne '*') {
			$Excel.Workbook.Worksheets[$WorksheetName]
		} else {
			$Excel.Workbook.Worksheets
		}

		foreach ($ws in $Worksheets) {
			if ($ExcludeHiddenSheet -and $ws.Hidden -ne 'visible') {continue}
			$Tables = if ($tabid.count) {
				$ws.Tables[$tabid]
			} else {
				$ws.Tables
			}

			$tabcollection = [ordered]@{
				WorksheetName = $ws.name
				Tables        = [ordered]@{}
			}

			if ($Content) {
				foreach ($Table in $Tables) {
					if ([string]::IsNullOrEmpty($Table.name)) {continue}
					##if (-not $Table.Address.Address) {continue}
					$rowCount      = $Table.Address.Rows
					$colCount      = $Table.Address.Columns
					$start,$end    = $Table.Address.Address.Split(':')
					$pos           = $start.IndexOfAny('0123456789'.ToCharArray())
					[int]$startCol = ConvertFrom-ExcelColumnName $start.Substring(0,$pos)
					[int]$startRow = $start.Substring($pos)
					$tabwidth      = $startCol + $colCount # relative table width - horisontal border
					
					# Table header
					$propertyNames = for ($col=$startCol; $col -lt $tabwidth; $col++) {
						$ws.Cells[$startRow, $col].value
					}

					$tabheight = $startRow + $rowCount + 1 # relative table height - vertical border
					$tabcollection['Tables'][$Table.name] = for ($row=($startRow+1); $row -lt $tabheight; $row++) {
						$nextrow = [ordered]@{}
						for (($col=$startCol),($i=0); $col -lt $tabwidth; $col++,$i++) {
							$nextrow.($propertyNames[$i]) = $ws.Cells[$row, $col].value
						}
						[PSCustomObject]$nextrow
					} # rows
				} # table contents
			} 
			else {
				$tabcollection['Tables'] = @($Tables.Name).where{-not [string]::IsNullOrEmpty($_)}
				##$tabcollection['Tables'] = @($Tables).where{$_.Address.Address}.foreach{$_.Name}
			}

			# TODO: extract "visual" tables from grid
			#if ($GridTable -and $Content -and -not $tabcollection['Tables'].count) {}
			if ($tabcollection['Tables'].count -or $IncludeEmptySheet) {
				[PSCustomObject]$tabcollection
			}
		} # sheets

		if (-not $ExcelPackage) {
			$Stream.Close()
			$Stream.Dispose()
		}
		$Excel.Dispose()
		$Excel = $null
	} # end
} # END Get-ExcelTable