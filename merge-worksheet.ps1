Function Merge-Worksheet {
   <#
      .Synopsis
        Merges two worksheets (or other objects) into a single worksheet with differences marked up.
      .Description
        The Compare-Worksheet command takes two worksheets and marks differences in the source document, and optionally outputs a grid showing the changes.
        By contrast the Merge-Worksheet command takes the worksheets and combines them into a single sheet showing the old and new data side by side .
        Although it is designed to work with Excel data it can work with arrays of any kind of object; so it can be a merge *of* worksheets, or a merge *to* worksheet. 
      .Example 
        merge-worksheet "Server54.xlsx" "Server55.xlsx" -WorkSheetName services -OutputFile Services.xlsx -OutputSheetName 54-55 -show
        The workbooks contain audit information for two servers, one page contains a list of services. This command creates a worksheet named 54-55 
        in a workbook named services and shows all the services and their differences, and opens it in Excel 
      .Example 
        merge-worksheet "Server54.xlsx" "Server55.xlsx" -WorkSheetName services -OutputFile Services.xlsx -OutputSheetName 54-55 -HideEqual -AddBackgroundColor LightBlue -show
        This modifies the previous command to hide the equal rows in the output sheet and changes the color used to mark rows "Added" to the second file.  
      .Example
        merge-worksheet -OutputFile .\j1.xlsx -OutputSheetName test11 -ReferenceObject (dir .\ImportExcel\4.0.7) -DifferenceObject (dir '\ImportExcel\4.0.8') -Property Length -Show
        This version compares two directories, and marks what has changed. 
        Because no "Key" property is given, "Name" is assumed to be the key and the only other property examined is length.  
        Files which are added or deleted or have changedd size will be highlighed in the output sheet. Changes to dates or other attributes will be ignored
      .Example
        merge-worksheet -Outf .\dummy.xlsx  -RefO (dir .\ImportExcel\4.0.7) -DiffO (dir .\ImportExcel\4.0.8') -Pr Length  -WhatIf -Passthru | Out-GridView 
        This time no file is written because -WhatIf is specified, and -Passthru causes the results to go Out-Gridview. This version uses aliases to shorten the parameters, 
        (OutputFileName can be "outFile" and the sheet "OutSheet" :  DifferenceObject & RefeenceObject can be DiffObject & RefObject)      
   #>
   [cmdletbinding(SupportsShouldProcess=$true)] 
   Param(
        #First Excel file to compare. You can compare two Excel files or two other objects but not one of each.
        [parameter(ParameterSetName='A',Mandatory=$true,Position=0)]
        [parameter(ParameterSetName='B',Mandatory=$true,Position=0)]
        [parameter(ParameterSetName='C',Mandatory=$true,Position=0)]
        $Referencefile ,

        #Second Excel file to compare.
        [parameter(ParameterSetName='A',Mandatory=$true,Position=1)]
        [parameter(ParameterSetName='B',Mandatory=$true,Position=1)]
        [parameter(ParameterSetName='C',Mandatory=$true,Position=1)]
        $Differencefile , 

        #Name(s) of worksheets to compare,
        [parameter(ParameterSetName='A',Position=2)]
        [parameter(ParameterSetName='B',Position=2)]
        [parameter(ParameterSetName='C',Position=2)]
        $WorkSheetName   = "Sheet1",
        
        #The row from where we start to import data, all rows above the StartRow are disregarded. By default this is the first row.
        [parameter(ParameterSetName='A')]
        [parameter(ParameterSetName='B')]
        [parameter(ParameterSetName='C')]
        [int]$Startrow = 1, 
        
        #Specifies custom property names to use, instead of the values defined in the column headers of the TopRow.
        [Parameter(ParameterSetName='B',Mandatory=$true)]
        [String[]]$Headername,   

        #Automatically generate property names (P1, P2, P3, ..) instead of the using the values the top row of the sheet.
        [Parameter(ParameterSetName='C',Mandatory=$true)]
        [switch]$NoHeader, 
       
        [parameter(ParameterSetName='D',Mandatory=$true)]
        [Alias('RefObject')]
        $ReferenceObject ,
        [parameter(ParameterSetName='D',Mandatory=$true)]
        [Alias('DiffObject')]
        $DifferenceObject ,

        #File to hold merged data.
        [parameter(Mandatory=$true,Position=3)]
        [Alias('OutFile')]
        $OutputFile , 
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
        [System.Drawing.Color]$KeyFontColor          = "DarkRed", 
        #Sets the background color for changed rows. 
        [System.Drawing.Color]$ChangeBackgroundColor = "Orange",
        #Sets the background color for rows in the reference but deleted from the difference sheet. 
        [System.Drawing.Color]$DeleteBackgroundColor = "LightPink", 
        #Sets the background color for rows not in the reference but added to the difference sheet. 
        [System.Drawing.Color]$AddBackgroundColor    = "PaleGreen",   
        #if Specified hides the rows in the spreadsheet that are equal and only shows changes, added or deleted rows. 
        [switch]$HideEqual ,
        #If specified outputs the data to the pipeline (you can add -whatif so it the command only outputs to the command)
        [switch]$Passthru  ,
        #If specified, opens the output workbook.  
        [Switch]$Show
   )
    
#region Read Excel data
    if ($Referencefile -and $Differencefile) {
        #if the filenames don't resolve, give up now. 
        try     { $oneFile = ((Resolve-Path -Path $Referencefile -ErrorAction Stop).path -eq (Resolve-Path -Path $Differencefile  -ErrorAction Stop).path)}
        Catch   { Write-Warning -Message "Could not Resolve the filenames." ; return } 
        
        #If we have one file , we must have two different worksheet names. If we have two files worksheet can be a single string or two strings. 
        if      ($onefile -and ( ($WorkSheetName.count -ne 2) -or $WorkSheetName[0] -eq $WorkSheetName[1] ) ) {
            Write-Warning -Message "If both the Reference and difference file are the same then worksheet name must provide 2 different names" 
            return
        }
        if      ($WorkSheetName.count -eq 2)  {$worksheet1      = $WorkSheetName[0]  ;  $WorkSheet2 = $WorkSheetName[1]} 
        elseif  ($WorkSheetName -is [string]) {$worksheet1      = $WorkSheet2 = $WorkSheetName}
        else    {Write-Warning -Message "You must provide either a single worksheet name or two names." ; return }   
    
        $params= @{ ErrorAction = [System.Management.Automation.ActionPreference]::Stop } 
        foreach ($p in @("HeaderName","NoHeader","StartRow")) {if ($PSBoundParameters[$p]) {$params[$p] = $PSBoundParameters[$p]}}
        try     {
            $ReferenceObject  = Import-Excel -Path $Referencefile  -WorksheetName $WorkSheet1 @params                                                                       
            $DifferenceObject = Import-Excel -Path $Differencefile -WorksheetName $WorkSheet2 @Params 
        }
        Catch   {Write-Warning -Message "Could not read the worksheet from $Referencefile and/or $Differencefile." ; return } 
        $firstDataRow    =  + 1 ; 
        if ($NoHeader) {$firstDataRow = $Startrow  } else {$firstDataRow = $Startrow + 1}   
    }
    else {$firstDataRow = 1 } 
#endregion    

#region Set lists of properties and row numbers
    #Make a list of properties/headings using the Property (default "*") and ExcludeProperty parameters 
    $propList     = @() 
    $headings     = $ReferenceObject[-1].psobject.Properties.Name # This preserves the sequence - using get-member would sort them alphabetically!
    if ($NoHeader -and "Name" -eq $Key) {$Key      = "p1"} 
    if ($headings -notcontains    $Key) {Write-Warning -Message "You need to specify one of the headings in the sheet '$worksheet1' as a key." ; return }
    foreach ($p in $Property)           { $propList += ($headings.where({$_ -like    $p}) )} 
    foreach ($p in $ExcludeProperty)    { $propList  =  $propList.where({$_ -notlike $p})  } 
    if ($propList -notcontains $Key)    { $propList +=  $Key}    #If $key isn't one of the headings we will have bailed by now 
    $propList     = $propList   | Select-Object -Unique      #so, prolist must contain at least $key if nothing else 
    #Build the list of the properties to output, in order.  
    $diffpart     = @() 
    $refpart      = @()
    foreach ($p in $proplist.Where({$key -ne $_}) ) {$refPart += "<=$p" ; $diffPart += "=>$p" } 
    $outputProps  = @($key) + $refpart + $diffpart 
    #Key will go in column A, last reference column will be A if there is one property, B  if there are two, C if theere are 3 etc
    $lastRefCol   = [char](64 +       $propList.count)
    #First difference column will be the next one (we'll trap the case of only having the key later)  
    $FirstDiffCol = [char](65 +       $propList.count)
    #Last difference column will be A if there is one property, C if there are two, E if there are 3 
    $lastDiffCol  = [char](64 +  2  * $propList.count)
    
    #Add RowNumber to every row
    #If one sheet has extra rows we can get a single "==" result from compare, but with the row from the reference sheet 
    #but the row in the other sheet might so we will look up the row number from the key field build a hash table for that  
    
    $Rowhash = @{}  
    $i = $firstDataRow  ;  foreach ($row in $ReferenceObject)  {Add-Member -InputObject $row -MemberType NoteProperty -Name "_Row" -Value ($i ++) -Force} 
    $i = $firstDataRow  ;  foreach ($row in $DifferenceObject) {Add-Member -InputObject $row -MemberType NoteProperty -Name "_Row" -Value ($i ++) -Force
                                                               $Rowhash[$row.$key] = $row._row          
    } 
#endregion 

    $expandedDiff = Compare-Object -ReferenceObject $ReferenceObject -DifferenceObject $DifferenceObject -Property $propList -PassThru -IncludeEqual | 
                        Group-Object -Property $key | ForEach-Object {  
                            #The value of the key column is the name of the group. 
                            $keyval = $_.name 
                            #we're going to create a custom object from a hash table. We want the fields to be ordered
                            $hash = [ordered]@{}  
                            foreach ($result in $_.Group) {
                                if     ($result.SideIndicator -ne "=>")      {$hash["<Row"] = $result._Row  } 
                                elseif (-not $hash["<Row"])                  {$hash["<Row"] = "" }
                                #if we have already set the side, this is the second record, so set side to indicate "changed"     
                                if     ($hash.Side) {$hash.Side = "<>"} else {$hash["Side"] = $result.SideIndicator}
                                #if result is "in reference" and we don't have a matching "in difference" (meaning a change) the lookup will be blank. Which we want.   
                                $hash[">Row"] = $Rowhash[$keyval] 
                                $hash[$key]   = $keyval  
                                #For all the other fields we care about, create <=FieldName and/or =>FieldName 
                                foreach ($p in $propList.Where({$_ -ne $key})) {
                                    if  ($result.SideIndicator -eq "==")  {$hash[("=>$P")] = $hash[("<=$P")] =$result.$P}
                                    else                                  {$hash[($result.SideIndicator+$P)] =$result.$P}
                                }
                            } 
                            [Pscustomobject]$hash
    } | Sort-Object -Property "<row"  
    
    #Already sorted by reference row number, and fill in any blanks in the differene-row column
    for ($i = 1; $i -lt $expandedDiff.Count; $i++) {if (-not $expandedDiff[$i].">row") {$expandedDiff[$i].">row" = $expandedDiff[$i-1].">row" } }   
    #Sort by difference row number, and fill in any blanks in the reference-row column
    $expandedDiff = $expandedDiff | Sort-Object -Property ">row"  
    for ($i = 1; $i -lt $expandedDiff.Count; $i++) {if (-not $expandedDiff[$i]."<row") {$expandedDiff[$i]."<row" = $expandedDiff[$i-1]."<row" } }  
    if ($PSCmdlet.ShouldProcess($OutputFile,"Write Output to Excel file")) {            
        $xl = $expandedDiff  | Select-Object -Property $outputProps |  
                 Export-Excel -Path $OutputFile -WorkSheetname $OutputSheetName -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -PassThru 
        $ws =  $xl.Workbook.Worksheets[$OutputSheetName]   
        for ($i = 0; $i -lt $expandedDiff.Count; $i++ ) {
            if     ( $expandedDiff[$i].side -ne "==" )  {
                Set-Format -WorkSheet $ws -Range ("A" + ($i + 2 )) -FontColor       $KeyFontColor
            }
            elseif ( $HideEqual                      )  {$ws.row($i+2).hidden = $true }
            if     ( $expandedDiff[$i].side -eq "<>" )  {
                $range = $ws.Dimension -replace "\d+",  ($i + 2 )
                Set-Format -WorkSheet $ws -Range $range            -BackgroundColor $ChangeBackgroundColor
            }
            elseif ( $expandedDiff[$i].side -eq "<=" )  {
                $range = "A" + ($i + 2 ) + ":" + $lastRefCol + ($i + 2 ) 
                Set-Format -WorkSheet $ws -Range $range            -BackgroundColor $DeleteBackgroundColor 
            }
            elseif ( $expandedDiff[$i].side -eq "=>" )  {
                if ($propList.count -gt 1) {
                    $range = $FirstDiffCol + ($i + 2 ) + ":" + $lastDiffCol + ($i + 2 ) 
                    Set-Format -WorkSheet $ws -Range $range        -BackgroundColor $AddBackgroundColor
                }
                Set-Format -WorkSheet $ws -Range ("A" + ($i + 2 )) -BackgroundColor $AddBackgroundColor  
            }     
        }
        Close-ExcelPackage -ExcelPackage $xl -Show:$Show  
    }
    if ($PassThru) {$expandedDiff} 
}