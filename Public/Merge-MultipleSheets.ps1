function Merge-MultipleSheets {
     [CmdletBinding()]
     [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification="False positives when initializing variable in begin block")]
     [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification="MultipleSheet would be incorrect")]
     #[Alias("Merge-MulipleSheets")] #There was a spelling error in the first release. This was there to ensure things didn't break but intelisense gave the alias first.
     param   (
         [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
         $Path  ,
         [int]$Startrow = 1,

         [String[]]$Headername,

         [switch]$NoHeader,

         $WorksheetName   = "Sheet1",
         [Alias('OutFile')]
         $OutputFile = ".\temp.xlsx",
         [Alias('OutSheet')]
         $OutputSheetName = "Sheet1",
         $Property        = "*"    ,
         $ExcludeProperty ,
         $Key           = "Name" ,
         $KeyFontColor          = [System.Drawing.Color]::Red,
         $ChangeBackgroundColor = [System.Drawing.Color]::Orange,
         $DeleteBackgroundColor = [System.Drawing.Color]::LightPink,
         $AddBackgroundColor    = [System.Drawing.Color]::Orange,
         [switch]$HideRowNumbers ,
         [switch]$Passthru  ,
         [Switch]$Show
    )
     begin   {    $filestoProcess   = @()  }
     process {    $filestoProcess  += $Path}
     end     {
         if     ($filestoProcess.Count -eq 1 -and $WorksheetName -match '\*') {
             Write-Progress -Activity "Merging sheets" -CurrentOperation "Expanding * to names of sheets in $($filestoProcess[0]). "
             $excel = Open-ExcelPackage -Path $filestoProcess
             $WorksheetName = $excel.Workbook.Worksheets.Name.where({$_ -like $WorksheetName})
             Close-ExcelPackage -NoSave -ExcelPackage $excel
         }

         #Merge identically named sheets in different work books;
          if     ($filestoProcess.Count -ge 2 -and $WorksheetName -is "string" ) {
             Get-Variable -Name 'HeaderName','NoHeader','StartRow','Key','Property','ExcludeProperty','WorksheetName' -ErrorAction SilentlyContinue |
                 Where-Object {$_.Value} | ForEach-Object -Begin {$params= @{} } -Process {$params[$_.Name] = $_.Value}

             Write-Progress -Activity "Merging sheets" -CurrentOperation "comparing '$WorksheetName' in $($filestoProcess[-1]) against $($filestoProcess[0]). "
             $merged            = Merge-Worksheet  @params -Referencefile $filestoProcess[0] -Differencefile $filestoProcess[-1]
             $nextFileNo        = 2
             while ($nextFileNo -lt $filestoProcess.count -and $merged) {
                 Write-Progress -Activity "Merging sheets" -CurrentOperation "comparing '$WorksheetName' in $($filestoProcess[-$nextFileNo]) against $($filestoProcess[0]). "
                 $merged        = Merge-Worksheet  @params -ReferenceObject $merged -Differencefile $filestoProcess[-$nextFileNo]
                 $nextFileNo    ++

             }
         }
         #Merge different sheets from one workbook
         elseif ($filestoProcess.Count -eq 1 -and $WorksheetName.Count -ge 2 ) {
             Get-Variable -Name 'HeaderName','NoHeader','StartRow','Key','Property','ExcludeProperty' -ErrorAction SilentlyContinue |
                 Where-Object {$_.Value} | ForEach-Object -Begin {$params= @{} } -Process {$params[$_.Name] = $_.Value}

             Write-Progress -Activity "Merging sheets" -CurrentOperation "Comparing $($WorksheetName[-1]) against $($WorksheetName[0]). "
             $merged          = Merge-Worksheet  @params -Referencefile $filestoProcess[0] -Differencefile $filestoProcess[0] -WorksheetName $WorksheetName[0,-1]
             $nextSheetNo     = 2
             while ($nextSheetNo -lt $WorksheetName.count -and $merged) {
                 Write-Progress -Activity "Merging sheets" -CurrentOperation "Comparing $($WorksheetName[-$nextSheetNo]) against $($WorksheetName[0]). "
                 $merged      = Merge-Worksheet  @params -ReferenceObject $merged -Differencefile $filestoProcess[0] -WorksheetName  $WorksheetName[-$nextSheetNo] -DiffPrefix $WorksheetName[-$nextSheetNo]
                 $nextSheetNo ++
             }
         }
         #We either need one Worksheet name and many files or one file and many sheets.
         else {            Write-Warning -Message "Need at least two files to process"           ; return }
         #if the process didn't return data then abandon now.
         if (-not $merged) {Write-Warning -Message "The merge operation did not return any data."; return }

         $orderByProperties  = $merged[0].psobject.properties.where({$_.name -match "row$"}).name
         Write-Progress -Activity "Merging sheets" -CurrentOperation "creating output sheet '$OutputSheetName' in $OutputFile"
         $excel                 = $merged | Sort-Object -Property $orderByProperties  |
                                   Export-Excel -Path $OutputFile -WorksheetName $OutputSheetName -ClearSheet -BoldTopRow -AutoFilter -PassThru
         $sheet                 = $excel.Workbook.Worksheets[$OutputSheetName]

         #We will put in a conditional format for "if all the others are not flagged as 'same'" to mark rows where something is added, removed or changed
         $sameChecks            = @()

         #All the 'difference' columns in the sheet are labeled with the file they came from, 'reference' columns need their
         #headers prefixed with the ref file name,  $colnames is the basis of a regular expression to identify what should have $refPrefix appended
         $colNames              = @("^_Row$")
         if ($Key -ne "*")
               {$colnames      += "^$Key$"}
         if ($filesToProcess.Count -ge 2) {
               $refPrefix       = (Split-Path -Path $filestoProcess[0] -Leaf) -replace "\.xlsx$"," "
         }
         else {$refPrefix       = $WorksheetName[0] + " "}
         Write-Progress -Activity "Merging sheets" -CurrentOperation "applying formatting to sheet '$OutputSheetName' in $OutputFile"
         #Find the column headings which are in the form "diffFile  is"; which will hold 'Same', 'Added' or 'Changed'
         foreach ($cell in $sheet.Cells[($sheet.Dimension.Address -replace "\d+$","1")].Where({$_.value -match "\sIS$"}) ) {
            #Work leftwards across the headings applying conditional formatting which says
            # 'Format this cell if the "IS" column has a value of ...' until you find a heading which doesn't have the prefix.
            $prefix             = $cell.value -replace  "\sIS$",""
            $columnNo           = $cell.start.Column -1
            $cellAddr           = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R1C$columnNo",1,$columnNo)
            while ($sheet.cells[$cellAddr].value -match $prefix) {
                $condFormattingParams =  @{RuleType='Expression'; BackgroundPattern='Solid'; Worksheet=$sheet; StopIfTrue=$true; Range=$([OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R[1]C[$columnNo]:R[1048576]C[$columnNo]",0,0)) }
                Add-ConditionalFormatting @condFormattingParams -ConditionValue ($cell.Address + '="Added"'  ) -BackgroundColor $AddBackgroundColor
                Add-ConditionalFormatting @condFormattingParams -ConditionValue ($cell.Address + '="Changed"') -BackgroundColor $ChangeBackgroundColor
                Add-ConditionalFormatting @condFormattingParams -ConditionValue ($cell.Address + '="Removed"') -BackgroundColor $DeleteBackgroundColor
                $columnNo --
                $cellAddr       = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R1C$columnNo",1,$columnNo)
            }
            #build up a list of prefixes in $colnames - we'll use that to set headers on rows from the reference file; and build up the "if the 'is' cell isn't same" list
            $colNames          += $prefix
            $sameChecks        += (($cell.Address -replace "1","2") +'<>"Same"')
         }

         #For all the columns which don't match one of the Diff-file prefixes or "_Row" or the 'Key' columnn name; add the reference file prefix to their header.
         $nameRegex             = $colNames -Join '|'
         foreach ($cell in $sheet.Cells[($sheet.Dimension.Address -replace "\d+$","1")].Where({$_.value -Notmatch $nameRegex}) ) {
            $cell.Value         = $refPrefix + $cell.Value
            $condFormattingParams =  @{RuleType='Expression'; BackgroundPattern='Solid'; Worksheet=$sheet; StopIfTrue=$true; Range=[OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R[2]C[$($cell.start.column)]:R[1048576]C[$($cell.start.column)]",0,0)}
            Add-ConditionalFormatting @condFormattingParams -ConditionValue ("OR("  +(($sameChecks -join ",") -replace '<>"Same"','="Added"'  ) +")" )   -BackgroundColor $DeleteBackgroundColor
            Add-ConditionalFormatting @condFormattingParams -ConditionValue ("AND(" +(($sameChecks -join ",") -replace '<>"Same"','="Changed"') +")" )   -BackgroundColor $ChangeBackgroundColor
         }
         #We've made a bunch of things wider so now is the time to autofit columns. Any hiding has to come AFTER this, because it unhides things
         if ($env:NoAutoSize) {Write-Warning "Autofit is not available with this OS configuration."}
         else  {$sheet.Cells.AutoFitColumns()}

         #if we have a key field (we didn't concatenate all fields) use what we built up in $sameChecks to apply conditional formatting to it (Row no will be in column A, Key in Column B)
         if ($Key -ne '*') {
               Add-ConditionalFormatting -Worksheet $sheet -Range "B2:B1048576" -ForeGroundColor $KeyFontColor -BackgroundPattern 'None' -RuleType Expression -ConditionValue ("OR(" +($sameChecks -join ",") +")" )
               $sheet.view.FreezePanes(2, 3)
         }
         else {$sheet.view.FreezePanes(2, 2) }
         #Go back over the headings to find and hide the "is" columns;
         foreach ($cell in $sheet.Cells[($sheet.Dimension.Address -replace "\d+$","1")].Where({$_.value -match "\sIS$"}) ) {
            $sheet.Column($cell.start.Column).HIDDEN = $true
         }

         #If specified, look over the headings for "row" and hide the columns which say "this was in row such-and-such"
         if ($HideRowNumbers) {
            foreach ($cell in $sheet.Cells[($sheet.Dimension.Address -replace "\d+$","1")].Where({$_.value -match "Row$"}) ) {
                $sheet.Column($cell.start.Column).HIDDEN = $true
            }
         }
         if ($Passthru) {$excel}
         else {Close-ExcelPackage -ExcelPackage $excel -Show:$Show}
         Write-Progress -Activity "Merging sheets" -Completed
     }
 }
