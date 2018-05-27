    Param   (
         [Parameter(Mandatory=$true,ValueFromPipeline=$true)]  
         [string[]]$Path  ,
         #The row from where we start to import data, all rows above the StartRow are disregarded. By default this is the first row.
         [int]$Startrow = 1, 
         
         #Specifies custom property names to use, instead of the values defined in the column headers of the TopRow.
         [String[]]$Headername,   
 
         #Automatically generate property names (P1, P2, P3, ..) instead of the using the values the top row of the sheet.
         [switch]$NoHeader, 
        
         #Name(s) of worksheets to compare,
         $WorkSheetName   = "Sheet1",
         #File to write output to 
         [Alias('OutFile')]
         $OutputFile = ".\temp.xlsx", 
         #Name of worksheet to output - if none specified will use the reference worksheet name. 
         [Alias('OutSheet')]
         $OutputSheetName = "Sheet1",
         #Properties to include in the DIFF - supports wildcards, default is "*".
         $Property        = "*"    ,
         #Properties to exclude from the the search - supports wildcards. 
         $ExcludeProperty ,
         #Name of a column which is unique used to pair up rows from the refence and difference side, default is "Name".
         $Key           = "Name" ,
         #Sets the font color for the "key" field; this means you can filter by color to get only changed rows. 
         [System.Drawing.Color]$KeyFontColor          = "Red", 
         #Sets the background color for changed rows. 
         [System.Drawing.Color]$ChangeBackgroundColor = "Orange",
         #Sets the background color for rows in the reference but deleted from the difference sheet. 
         [System.Drawing.Color]$DeleteBackgroundColor = "LightPink", 
         #Sets the background color for rows not in the reference but added to the difference sheet. 
         [System.Drawing.Color]$AddBackgroundColor    = "Orange",   
         #if Specified hides the columns in the spreadsheet that contain the row numbers 
         [switch]$HideRowNumbers ,
         #If specified outputs the data to the pipeline (you can add -whatif so it the command only outputs to the command)
         [switch]$Passthru  ,
         #If specified, opens the output workbook.  
         [Switch]$Show
    )
    begin   { $filestoProcess  = @()  }
    process { $filestoProcess += $Path} 
    end     {
        if  ( $filestoProcess.count -lt 2) {Write-Warning -Message "Need at least two files to process"; return} 
        
        #Set up the parameters we will pass to merge worksheet 
        Get-Variable -Name 'HeaderName','NoHeader','StartRow','Key','Property','ExcludeProperty','WorkSheetName' -ErrorAction SilentlyContinue |
            Where-Object {$_.Value} | ForEach-Object -Begin {$params= @{} } -Process {$params[$_.Name] = $_.Value} 
        
        Write-Progress -Activity "Merging sheets" -CurrentOperation "Comparing $($filestoProcess[-1]) against $($filestoProcess[0]). "
        $merged         = Merge-Worksheet  @params -Referencefile $filestoProcess[0] -Differencefile $filestoProcess[-1]
        $nextFileNo     = 2
        while ($nextFileNo -lt $filestoProcess.count -and $merged) { 
            Write-Progress -Activity "Merging sheets" -CurrentOperation "Comparing $($filestoProcess[-$nextFileNo]) against $($filestoProcess[0]). "
            $merged     = Merge-Worksheet  @params -ReferenceObject $merged -Differencefile $filestoProcess[-$nextFileNo]     
            $nextFileNo ++
        }
        if (-not $merged) {Write-Warning -Message "The merge operation did not return any data."; return }

        Write-Progress -Activity "Merging sheets" -CurrentOperation "Creating output sheet '$OutputSheetName' in $OutputFile"
        $excel          = $merged | Sort-Object "_row"  | Update-FirstObjectProperties | Export-Excel -Path $OutputFile -WorkSheetname $OutputSheetName -ClearSheet -FreezeTopRow -BoldTopRow -AutoFilter -PassThru 
        $sheet          = $excel.Workbook.Worksheets[$OutputSheetName]   
        
        #We will put in a conditional format for "if all the others are not flagged as 'same'" to mark rows where something is added, removed or changed
        $sameChecks    = @()
        
        #All the 'difference' columns in the sheet are labeled with the file they came from, 'reference' columns need their 
        #headers prefixed with the ref file name,  $colnames is the basis of a regular expression to identify what should have $refPrefix appended  
        $colNames      = @("_Row","^$Key`$") 
        $refPrefix     = (Split-Path -Path $filestoProcess[0] -Leaf) -replace "\.xlsx$"," "  
        
        Write-Progress -Activity "Merging sheets" -CurrentOperation "Applying formatting to sheet '$OutputSheetName' in $OutputFile"
        #Find the column headings which are in the form "diffFile  is"; which will hold 'Same', 'Added' or 'Changed'  
        foreach ($cell in $sheet.Cells[($sheet.Dimension.Address -replace "\d+$","1")].Where({$_.value -match "\sIS$"}) ) {
            #Work leftwards across the headings applying conditional formatting which says 
            # 'Format this cell if the "IS" column has a value of ...' until you find a heading which doesn't have the prefix. 
            $prefix    = $cell.value -replace  "\sIS$","" 
            $columnNo  = $cell.start.Column -1 
            $cellAddr  = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R1C$columnNo",1,$columnNo) 
            while ($sheet.cells[$cellAddr].value -match $prefix) {
                $condFormattingParams =  @{RuleType='Expression'; BackgroundPattern='None'; WorkSheet=$sheet; Range=$([OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R[1]C[$columnNo]:R[1048576]C[$columnNo]",0,0)) }    
                Add-ConditionalFormatting @condFormattingParams -ConditionValue ($cell.Address + '="Added"'  ) -BackgroundColor $AddBackgroundColor 
                Add-ConditionalFormatting @condFormattingParams -ConditionValue ($cell.Address + '="Changed"') -BackgroundColor $ChangeBackgroundColor 
                Add-ConditionalFormatting @condFormattingParams -ConditionValue ($cell.Address + '="Removed"') -BackgroundColor $DeleteBackgroundColor 
                $columnNo -- 
                $cellAddr = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R1C$columnNo",1,$columnNo) 
            }
            #build up a list of prefixes in $colnames - we'll use that to set headers on rows from the reference file; and build up the "if the 'is' cell isn't same" list
            $colNames    += $prefix
            $sameChecks  += (($cell.Address -replace "1","2") +'<>"Same"')
        }
        
        #For all the columns which don't match one of the Diff-file prefixes or "_Row" or the 'Key' columnn name; add the reference file prefix to their header.   
        $nameRegex = $colNames -Join "|"  
        foreach ($cell in $sheet.Cells[($sheet.Dimension.Address -replace "\d+$","1")].Where({$_.value -Notmatch $nameRegex}) ) {
            $cell.Value = $refPrefix + $cell.Value
        }
        #We've made a bunch of things wider so now is the time to autofit columns. Any hiding has to come AFTER this, because it unhides things 
        $sheet.Cells.AutoFitColumns() 
        
        #if we have a key field (we didn't concatenate all fields) use what we built up in $sameChecks to apply conditional formatting to it (Row no will be in column A, Key in Column B) 
        if ($Key -ne '*') {
            Add-ConditionalFormatting -WorkSheet $sheet -Range "B2:B1048576" -ForeGroundColor $KeyFontColor -BackgroundPattern 'None' -RuleType Expression -ConditionValue ("OR(" +($sameChecks -join ",") +")" )   
        }
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
        
        Close-ExcelPackage -ExcelPackage $excel -Show:$Show  
        Write-Progress -Activity "Merging sheets" -Completed
    }
