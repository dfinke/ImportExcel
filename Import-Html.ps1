function Import-Html {
	[CmdletBinding()]
	param(
		$url,
		$index,
		$Header,
		[ValidatePattern('xlsx')]
		[System.IO.FileInfo]$Path,
		[Switch]$Append,
		[int]$FirstDataRow = 0,
		[Switch]$UseDefaultCredentials,
		[Switch]$Show
	)
	
	if ($Path) {
		$TempFile = $False
		if ((Test-Path $Path) -and -not ($Append)) {
			
			$Message = 'File already exists, would you like to overwrite this file'
			$Question = 'Selecting Yes will overwrite file, No will exit'

			$Choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
			$Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
			$Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
		
			$Decision = $Host.UI.PromptForChoice($Message, $Question, $Choices, 1)
			if ($Decision -eq 0) {
				Remove-Item $Path -Force -Confirm:$False
			}
			else {
				break
			}
		}
	}
	else {
		$Path = [System.IO.Path]::GetTempFileName() -replace "tmp", "xlsx"
		Remove-Item $Path -ErrorAction Ignore
		$TempFile = $True
	}
	
	if ($Show) {
		$OpenExcel = $True
	}
	else {
		$OpenExcel = $False
		Write-Host "Exporting to Excel file $($Path)"
		if ($TempFile -eq $True) {
		Write-Verbose "Setting clipboard to $($Path)"
		$Path.FullName | clip
		}
	}
	
	if ($Append -and $TempFile -eq $False) {
		$AppendState = $True
	}
	else {
		$AppendState = $False
	}
	
	Write-Verbose "Exporting to Excel file $($Path)"

	Get-HtmlTable -url $url -tableIndex $index -Header $Header -FirstDataRow $FirstDataRow -UseDefaultCredentials: $UseDefaultCredentials | Export-Excel -Path $Path -AutoSize -Append:$AppendState -Show:$OpenExcel
}
