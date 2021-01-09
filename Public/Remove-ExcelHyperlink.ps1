function Remove-ExcelHyperlink {
    <#
            .SYNOPSIS
                Removes the hyperlink from cell(s) or worksheet(s) or entire workbook
    
            .PARAMETER Path
                The Excel workbook path
    
            .PARAMETER ExcelPackage
                The Excel package
         
            .PARAMETER WorksheetName
                The worksheet to remove hyperlink(s) from
     
            .PARAMETER Cell
                The cell to remove hyperlink from
            
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
        [String[]]$Cell,
        [switch]$Show
    )
    
        if (-not $WorksheetName -and $Cell) { Write-Warning -Message "Please provide the WorksheetName" ; return }

        Write-verbose -Message "Opening ExcelPackage via Path"
        if ($Path -and -not $ExcelPackage) { $ExcelPackage = Open-ExcelPackage -Path $Path }

        Write-verbose -Message "Getting ExcelSheetInfo"
        $WorksheetInfo = (Get-ExcelSheetInfo -Path $Path).Name
        if(!$WorksheetName) { 
            $WorksheetName = $WorksheetInfo
        }

        Write-verbose -Message "Looping through the worksheets"
        $cells = @()
        foreach ($worksheet in $WorksheetName) {
            if ($worksheet -notin $WorksheetInfo) { Write-Warning -Message "Worksheet [$worksheet] does not exist" ; continue }
            Write-verbose -Message "... worksheet: $worksheet"
            $ws = $ExcelPackage.Workbook.Worksheets[$worksheet]

            if(!$Cell) {
                Write-verbose -Message "No cell specified - checking all"
                $Cell = (Get-ExcelHyperlink -Path $path -WorksheetName $worksheet).Cell
            }
            foreach ($cellItem in $Cell) {

                if ($ws.Cells["$cellItem"].Hyperlink) {
                    $cellValue = $ws.Cells[$cellItem].Value
                    Write-verbose -Message "Removing hyperlink from [$cellItem] cell"
                    $ws.Cells["$cellItem"].Hyperlink = $null

                    Write-verbose -Message "Changing [$cellItem] cell style from [$($ws.Cells[$cellItem].StyleName)] to [Normal]"
                    $null = $ws.Cells[$cellItem].StyleName = 'Normal'
                    $ws.Cells[$cellItem].Value = $cellValue
                    }
                }
        }
        Write-verbose -Message "Closing the ExcelPackage"
        Close-ExcelPackage -ExcelPackage $ExcelPackage -Show:$Show      
}