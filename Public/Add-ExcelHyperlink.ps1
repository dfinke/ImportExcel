function Add-ExcelHyperlink {
    <#
            .SYNOPSIS
                Add the hyperlink in a cell (supports local named range only at the moment)
    
            .PARAMETER Path
                The Excel workbook path
    
            .PARAMETER ExcelPackage
                The Excel package
         
            .PARAMETER WorksheetName
                The worksheet containing the target cell
     
            .PARAMETER Cell
                Target cell that will have hyperlink set
         
            .PARAMETER DisplayName
                If specified will replace the value of the underlying cell
    
            .PARAMETER Show
                Display the workbook after adding the hyperlink
    
            .NOTES
                Author: Mikey Bronowski (@MikeyBronowski), bronowski.it
    
            .EXAMPLE		
                Remove-Item -Path $path -ErrorAction SilentlyContinue
                $excelPackage = 'Some text' | Export-Excel -Path $path -WorksheetName Sheet1 -PassThru
                $excel=$excelPackage.Workbook.Worksheets['Sheet1']
                $excelPackage.Workbook.Names.Add('NamedRange',$excel.cells['D1:F50'])
                Close-ExcelPackage $excelPackage
                Add-ExcelHyperlink -Path $path -WorksheetName Sheet1 -Hyperlink 'NamedRange'-Cell A1 -Show
    
                Add a hyperlink in A1 to NamedRange keeping cell's value as DisplayName
    
            .EXAMPLE		
                Remove-Item -Path $path -ErrorAction SilentlyContinue
                $excelPackage = 'Some text' | Export-Excel -Path $path -WorksheetName Sheet1 -PassThru
                $excel=$excelPackage.Workbook.Worksheets['Sheet1']
                $excelPackage.Workbook.Names.Add('NamedRange',$excel.cells['D1:F50'])
                Close-ExcelPackage $excelPackage
                Add-ExcelHyperlink -Path $path -WorksheetName Sheet1 -Hyperlink 'NamedRange'-Cell A2 -Show
    
                Add a hyperlink in A2 to NamedRange without setting DisplayName - it will be set to 'Link' by default
    
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
            [String]$WorksheetName,
            [Parameter(Mandatory)]
            [String]$Cell,
            [Parameter(Mandatory)]
            [String]$Hyperlink,
            [String]$DisplayName,
            [switch]$Show
        )
        
            Write-verbose -Message "Opening ExcelPackage via Path [$Path]"
            if ($Path -and -not $ExcelPackage) { $ExcelPackage = Open-ExcelPackage -Path $Path }
    
            Write-verbose -Message "Setting the Worksheet to [$WorksheetName]"
            $ws = $ExcelPackage.Workbook.Worksheets[$WorksheetName]
            
    
            $cellValue = $ws.Cells[$Cell].Value
    
            if (!$cellValue) {
                Write-verbose -Message "The [$Cell] cell did not have any values setting DisplayName to [Link]"
                $cellValue = 'Link'
            }
    
            if (!$DisplayName) {
                Write-verbose -Message "Keeping the value = [$cellValue] of the [$Cell] cell"
                $DisplayName = $cellValue
            }
    
            Write-verbose -Message "Creating a hyperlink [$Hyperlink] under [$DisplayName]"
            $hyperlinkObj = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $Hyperlink , $DisplayName
            
            Write-verbose -Message "Adding hyperlink [$Hyperlink] in the [$Cell] cell on [$WorksheetName] worksheet"
            $null = $ws.Cells[$Cell].Hyperlink = $hyperlinkObj
    
            if(($ws.Workbook.Styles.NamedStyles.Name -contains 'Hyperlink') -eq $false) {
                Write-verbose -Message "The NamedStyle Hyperlink does not exist - creating one"
                $namedStyle=$ws.Workbook.Styles.CreateNamedStyle('Hyperlink')
                $namedStyle.Style.Font.UnderLine = $true
                $namedStyle.Style.Font.Color.SetColor('Blue')
            }
    
            Write-verbose -Message "Changing [$Cell] cell style from [$($ws.Cells[$Cell].StyleName)] to [$($namedStyle.Name)]"
            $null = $ws.Cells[$Cell].StyleName = $namedStyle.Name
    
            Write-verbose -Message "Closing the ExcelPackage"
            Close-ExcelPackage -ExcelPackage $ExcelPackage -Show:$Show
    }