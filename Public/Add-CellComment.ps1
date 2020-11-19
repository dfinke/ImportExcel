function Add-CellComment {
<#
	.SYNOPSIS
		Adds a comment to the specified cell
 
	.PARAMETER Worksheet
		The worksheet containing the target cell
 
	.PARAMETER Column
		The maximum characters per line.
 
	.PARAMETER CellAddress
		The number of characters to indent each line.
 
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
			[string]$CellAddress,
		[Parameter(Mandatory = $true)]
			[string]$Comment,
		[string]$Author = "ImportExcel",
		[switch]$noautofit
	)

	$cellAddressPattern = [Regex]::new('[A-z]{1,2}[\d]+')
	if($($CellAddress -notmatch $cellAddressPattern))
	{
		Write-Error "Invalid cell specified"
		return
	}
	
	$cellComment = $Worksheet.Cells["$($Column)$($Row)"].AddComment($Comment,$author)

	if($noautofit -ne $true) {
		$cellComment.AutoFit = $true
	}
}