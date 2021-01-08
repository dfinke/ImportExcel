function Get-ExcelHyperlink {
    [CmdletBinding()]
    param(     
        [String]$Path,
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [String[]]$WorksheetName,
        [String[]]$Cell,
        [String]$Hyperlink,
        [String]$DisplayName,
        [switch]$Show
    )
    
        if (-not $WorksheetName -and $Cell) { Write-Warning -Message "Please provide the WorksheetName" ; return }

        Write-verbose -Message "Opening ExcelPackage via Path"
        if ($Path -and -not $ExcelPackage) { $ExcelPackage = Open-ExcelPackage -Path $Path }

        $WorksheetInfo = (Get-ExcelSheetInfo -Path $Path).Name

        Write-verbose -Message "Getting ExcelSheetInfo"
        if(!$WorksheetName) { 
            $WorksheetName = $WorksheetInfo
        }

        Write-verbose -Message "Looping through the ExcelSheets"
        $cells = @()
        foreach ($worksheet in $WorksheetName) {
            if ($worksheet -notin $WorksheetInfo) { Write-Warning -Message "Worksheet [$worksheet] does not exist" ; continue }
            Write-verbose -Message "Looping through the ExcelSheets: $worksheet"
            $ws = $ExcelPackage.Workbook.Worksheets[$worksheet]

            if($Cell) {Write-verbose -Message "Checking [$Cell] cell only"
                $ws.Cells["$Cell"] | SELECT Worksheet, Address, StyleName, Hyperlink
            }
            else {
                $ws.Cells | Where-Object {$_.Hyperlink -ne $null} | SELECT Worksheet, Address, StyleName, Hyperlink
            }
        }
        Close-ExcelPackage -ExcelPackage $ExcelPackage -NoSave -Show:$Show      
}