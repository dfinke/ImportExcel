function Get-ExcelHyperlink {
    <#
            .SYNOPSIS
                Get the hyperlink within workbook (supports local named range only at the moment)
    
            .PARAMETER Path
                The Excel workbook path
    
            .PARAMETER ExcelPackage
                The Excel package
         
            .PARAMETER WorksheetName
                The worksheet containing the target cell
     
            .PARAMETER Cell
                Target cell that will have hyperlink set
            
            .NOTES
                Author: Mikey Bronowski (@MikeyBronowski), bronowski.it
    
            .EXAMPLE		
                Get-ExcelHyperlink -Path $path
    
                Get all hyperlinks within Excel file
    
            .EXAMPLE		
                Get-ExcelHyperlink -Path $path -WorksheetName Sheet1

                Get all hyperlinks within worksheet

            .EXAMPLE		
                Get-ExcelHyperlink -Path $path -WorksheetName Sheet1 -Cell A2

                Get hyperlink details from the cell
    
            .EXAMPLE	
                Remove-Item -Path $path -ErrorAction SilentlyContinue
                $excelPackage = 'Some text' | Export-Excel -Path $path -WorksheetName Sheet1 -PassThru
                $excel=$excelPackage.Workbook.Worksheets['Sheet1']
                $excelPackage.Workbook.Names.Add('NamedRange',$excel.cells['D1:F50'])
                Close-ExcelPackage $excelPackage
                Add-ExcelHyperlink -Path $path -WorksheetName Sheet1 -Hyperlink 'NamedRange'-Cell A3 -DisplayName 'Link to NamedRange' -Show
    
                Add a hyperlink in A3 to NamedRange with a custom DisplayName
            #>
    [CmdletBinding()]
    param(     
        [String]$Path,
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [String[]]$WorksheetName,
        [String[]]$Cell
    )
    
        if (-not $WorksheetName -and $Cell) { Write-Warning -Message "Please provide the WorksheetName" ; return }

        Write-verbose -Message "Opening ExcelPackage via Path"
        if ($Path -and -not $ExcelPackage) { $ExcelPackage = Open-ExcelPackage -Path $Path }

        Write-verbose -Message "Getting ExcelSheetInfo"
        $WorksheetInfo = (Get-ExcelSheetInfo -Path $Path).Name
        if(!$WorksheetName) { 
            $WorksheetName = $WorksheetInfo
        }

        Write-verbose -Message "Looping through the ExcelSheets"
        $cells = @()
        foreach ($worksheet in $WorksheetName) {
            if ($worksheet -notin $WorksheetInfo) { Write-Warning -Message "Worksheet [$worksheet] does not exist" ; continue }
            Write-verbose -Message "Looping through the ExcelSheets: $worksheet"
            $ws = $ExcelPackage.Workbook.Worksheets[$worksheet]

            if($Cell) {
                foreach ($cellItem in $Cell) {
                    Write-verbose -Message "Checking [$cellItem] cell only"
                    $ws.Cells["$cellItem"] | SELECT Worksheet, Address, StyleName, Hyperlink
                }
            }
            else {
                $ws.Cells | Where-Object {$_.Hyperlink -ne $null} | SELECT Worksheet, @{N='Cell';E={$_.Address}}, StyleName, Hyperlink
            }
        }
        Write-verbose -Message "Closing the ExcelPackage"
        Close-ExcelPackage -ExcelPackage $ExcelPackage -NoSave -Show:$Show      
}