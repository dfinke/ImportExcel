function Get-ExcelCellComment {
	<#
		.SYNOPSIS
			Get the comment from the specified cell
	 
		.PARAMETER Worksheet
			The worksheet containing the target cell
	 
		.PARAMETER Column
			Column to add the comment to. This can either be a column letter/name
			or the index of the column
	 
		.PARAMETER Row
			Row to add the comment to

		.EXAMPLE
		
			Get-CellComment -Worksheet $excelPkg.Sheet1 -CellAddress A1
		#>
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[OfficeOpenXml.ExcelWorksheet]$Worksheet,
		[Parameter(Mandatory = $true)]
		[string]$Column,
		[Parameter(Mandatory = $true)]
		[int]$Row
	)

	# Convert column indexes to their names
	if ($Column -match "\d") {
		$Column = Get-ExcelColumnName -columnName $Column
	}

	$cellAddress = "$Column$Row"
	$cellAddressPattern = [Regex]::new('[A-z]{1,2}[\d]+')
	if ($($CellAddress -notmatch $cellAddressPattern)) {
		Write-Error "Invalid cell specified"
		return
	}

	# Comments are a collection, so not directly referencable by address
	$comment = $Worksheet.Comments | Where-Object {$_.Address -eq "$Column$Row"}

	return $comment
}
