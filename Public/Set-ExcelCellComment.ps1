function Set-ExcelCellComment {
	<#
		.SYNOPSIS
			Adds or updates a comment to the specified cell
	 
		.PARAMETER Worksheet
			The worksheet containing the target cell
	 
		.PARAMETER Column
			Column to add the comment to. This can either be a column letter/name
			or the index of the column
	 
		.PARAMETER Row
			Row to add the comment to
	 
		.PARAMETER Comment
			The comment to be added
	
		.PARAMETER Author
			The author of the comment, which is required for adding a comment, but 
			we provide a default value
	
		.PARAMETER noautofit
			If automatically resizing the comment is not desired that can be accomodated
	 
		.EXAMPLE
		
			Add-CellComment -Worksheet $excelPkg.Sheet1 -CellAddress A1 -Comment "This is a comment" -Author "Automated Process"
		#>
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[OfficeOpenXml.ExcelWorksheet]$Worksheet,
		[Parameter(Mandatory = $true)]
		[string]$Column,
		[Parameter(Mandatory = $true)]
		[int]$Row,
		[Parameter(Mandatory = $true)]
		[string]$Comment,
		[string]$Author,
		[switch]$noautofit
	)

	if($null -eq $Author) {
		$Author = ""
	}

	# Convert column indexes to their corresponding column names
	if ($Column -match "\d") {
		$Column = Get-ExcelColumnName -columnNumber $Column
	}

	$cellAddress = "$Column$Row"
	$cellAddressPattern = [Regex]::new('[A-z]{1,2}[\d]+')
	if ($($CellAddress -notmatch $cellAddressPattern)) {
		Write-Error "Invalid cell specified"
		return
	}
	
	# Check for an existing comment
	# Comments are a collection, so not directly referencable by address
	$cellComment = $Worksheet.Comments | Where-Object {$_.Address -eq "$Column$Row"}
	if($null -eq $cellComment) {
		$cellComment = $Worksheet.Cells[$CellAddress].AddComment($Comment, $author)
	}
	else {
		$cellComment.Text = $Comment
		$cellComment.Author = $Author
	}
	
	if ($noautofit -ne $true) {
		$cellComment.AutoFit = $true
	}
}
