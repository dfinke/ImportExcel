function Import-Html {
	[CmdletBinding()]
	param(
		$url,
		$index,
		$Header,
		[System.IO.FileInfo]$Path,
		[Switch]$Append,
		[int]$FirstDataRow = 0,
		[Switch]$UseDefaultCredentials,
		[Switch]$Show
	)
	
	if ($Path) {
		if ((Test-Path $Path) -and -not ($Append)) {
			Remove-Item $Path -Confirm
		}
	}
	else {
		$Path = [System.IO.Path]::GetTempFileName() -replace "tmp", "xlsx"
		Remove-Item $Path -ErrorAction Ignore	
	}
	
	if ($Show) {
		$OpenExcel = $True
	}
	else {
		$OpenExcel = $False
	}
	
	Write-Verbose "Exporting to Excel file $($Path)"

	$data = Get-HtmlTable -url $url -tableIndex $index -Header $Header -FirstDataRow $FirstDataRow -UseDefaultCredentials: $UseDefaultCredentials

	if ($Append -and $Path -notmatch 'temp') {$data | Export-Excel -Path $Path -AutoSize -Append -Show:$OpenExcel}
	else {$data | Export-Excel -Path $Path -AutoSize -Show:$OpenExcel}
}
