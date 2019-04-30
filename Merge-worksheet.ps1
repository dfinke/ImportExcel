Function Merge-Worksheet {
    <#
       .Synopsis
         Merges two Worksheets (or other objects) into a single Worksheet with differences marked up.
       .Description
         The Compare-Worksheet command takes two Worksheets and marks differences in the source document, and optionally outputs a grid showing the changes.
         By contrast the Merge-Worksheet command takes the Worksheets and combines them into a single sheet showing the old and new data side by side.
         Although it is designed to work with Excel data it can work with arrays of any kind of object; so it can be a merge *of* Worksheets, or a merge *to* a Worksheet.
       .Example
         Merge-Worksheet "Server54.xlsx" "Server55.xlsx" -WorksheetName services -OutputFile Services.xlsx -OutputSheetName 54-55 -show
         The workbooks contain audit information for two servers, one sheet contains
         a list of services. This command creates a worksheet named "54-55" in a
         workbook named "services.xlsx" which shows all the services and their
         differences, and opens the new workbook in Excel.
       .Example
         Merge-Worksheet "Server54.xlsx" "Server55.xlsx" -WorksheetName services -OutputFile Services.xlsx -OutputSheetName 54-55 -HideEqual -AddBackgroundColor LightBlue -show
         This modifies the previous command to hide the equal rows in the output
         sheet and changes the color used to mark rows added to the second file.
       .Example
         Merge-Worksheet -OutputFile .\j1.xlsx -OutputSheetName test11 -ReferenceObject (dir .\ImportExcel\4.0.7) -DifferenceObject (dir .\ImportExcel\4.0.8) -Property Length -Show
         This version compares two directories, and marks what has changed.
         Because no "Key" property is given, "Name" is assumed to be the key
         and the only other property examined is length. Files which are added
         or deleted or have changed size will be highlighed in the output sheet.
         Changes to dates or other attributes will be ignored.
       .Example
         Merge-Worksheet   -RefO (dir .\ImportExcel\4.0.7) -DiffO (dir .\ImportExcel\4.0.8) -Pr Length  | Out-GridView
         This time no file is written and the results - which include all properties,
         not just length, are output and sent to Out-Gridview. This version uses
         aliases to shorten the parameters, (OutputFileName can be "outFile" and
         the Sheet can be"OutSheet"; DifferenceObject & ReferenceObject can be
         DiffObject & RefObject respectively).
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    Param(
         #First Excel file to compare. You can compare two Excel files or two other objects or a reference obhct against a difference file, but not a reference file against an object.
         [parameter(ParameterSetName='A',Mandatory=$true,Position=0)]  #A = Compare two files default headers
         [parameter(ParameterSetName='B',Mandatory=$true,Position=0)]  #B = Compare two files user supplied headers
         [parameter(ParameterSetName='C',Mandatory=$true,Position=0)]  #C = Compare two files headers P1, P2, P3 etc
         $Referencefile ,

         #Second Excel file to compare.
         [parameter(ParameterSetName='A',Mandatory=$true,Position=1)]
         [parameter(ParameterSetName='B',Mandatory=$true,Position=1)]
         [parameter(ParameterSetName='C',Mandatory=$true,Position=1)]
         [parameter(ParameterSetName='E',Mandatory=$true,Position=1)] #D Compare two objects; E = Compare one object one file that uses default headers
         [parameter(ParameterSetName='F',Mandatory=$true,Position=1)] #F = Compare one object one file that uses user supplied headers
         [parameter(ParameterSetName='G',Mandatory=$true,Position=1)] #G   Compare one object one file that uses headers P1, P2, P3 etc
         $Differencefile ,

         #Name(s) of Worksheets to compare.
         [parameter(ParameterSetName='A',Position=2)]  #Applies to all sets EXCEPT D which is two objects (no sheets)
         [parameter(ParameterSetName='B',Position=2)]
         [parameter(ParameterSetName='C',Position=2)]
         [parameter(ParameterSetName='E',Position=2)]
         [parameter(ParameterSetName='F',Position=2)]
         [parameter(ParameterSetName='G',Position=2)]
         $WorksheetName   = "Sheet1",

         #The row from where we start to import data, all rows above the StartRow are disregarded. By default this is the first row.
         [parameter(ParameterSetName='A')]  #Applies to all sets EXCEPT D which is two objects (no sheets, so no start row )
         [parameter(ParameterSetName='B')]
         [parameter(ParameterSetName='C')]
         [parameter(ParameterSetName='E')]
         [parameter(ParameterSetName='F')]
         [parameter(ParameterSetName='G')]
         [int]$Startrow = 1,

         #Specifies custom property names to use, instead of the values defined in the column headers of the Start ROw.
         [Parameter(ParameterSetName='B',Mandatory=$true)]  #Compare  object + sheet or 2 sheets with user supplied headers
         [Parameter(ParameterSetName='F',Mandatory=$true)]
         [String[]]$Headername,

         #Automatically generate property names (P1, P2, P3, ..) instead of using the values the top row of the sheet.
         [Parameter(ParameterSetName='C',Mandatory=$true)]  #Compare  object + sheet or 2 sheets with headers of P1, P2, P3 ...
         [Parameter(ParameterSetName='G',Mandatory=$true)]
         [switch]$NoHeader,

         #Reference object to compare if a Worksheet is NOT being used. Reference object can combine with a difference sheet or difference object
         [parameter(ParameterSetName='D',Mandatory=$true)]
         [parameter(ParameterSetName='E',Mandatory=$true)]
         [parameter(ParameterSetName='F',Mandatory=$true)]
         [parameter(ParameterSetName='G',Mandatory=$true)]
         [Alias('RefObject')]
         $ReferenceObject ,
         #Difference object to compare if a Worksheet is NOT being used for either half. Can't have a reference sheet and difference object.
         [parameter(ParameterSetName='D',Mandatory=$true,Position=1)]
         [Alias('DiffObject')]
         $DifferenceObject ,
         [parameter(ParameterSetName='D',Position=2)]
         [parameter(ParameterSetName='E',Position=2)]
         [parameter(ParameterSetName='F',Position=2)]
         [parameter(ParameterSetName='G',Position=2)]
         #If there isn't a filename to use to label data from the "Difference" side, DiffPrefix is used, it defaults to "=>"
         $DiffPrefix = "=>" ,
         #File to hold merged data.
         [parameter(Position=3)]
         [Alias('OutFile')]
         $OutputFile ,
         #Name of Worksheet to output - if none specified will use the reference Worksheet name.
         [parameter(Position=4)]
         [Alias('OutSheet')]
         $OutputSheetName = "Sheet1",
         #Properties to include in the DIFF - supports wildcards, default is "*".
         $Property        = "*"    ,
         #Properties to exclude from the the search - supports wildcards.
         $ExcludeProperty ,
         #Name of a column which is unique used to pair up rows from the refence and difference side, default is "Name".
         $Key           = "Name" ,
         #Sets the font color for the "key" field; this means you can filter by color to get only changed rows.
         $KeyFontColor          = [System.Drawing.Color]::DarkRed ,
         #Sets the background color for changed rows.
         $ChangeBackgroundColor = [System.Drawing.Color]::Orange,
         #Sets the background color for rows in the reference but deleted from the difference sheet.
         $DeleteBackgroundColor = [System.Drawing.Color]::LightPink,
         #Sets the background color for rows not in the reference but added to the difference sheet.
         $AddBackgroundColor    = [System.Drawing.Color]::PaleGreen,
         #if specified, hides the rows in the spreadsheet that are equal and only shows changes, added or deleted rows.
         [switch]$HideEqual ,
         #If specified, outputs the data to the pipeline (you can add -WhatIf so the command only outputs to the pipeline).
         [switch]$Passthru  ,
         #If specified, opens the output workbook.
         [Switch]$Show
    )

 #region Read Excel data
    if ($Differencefile -is [System.IO.FileInfo]) {$Differencefile = $Differencefile.FullName}
    if ($Referencefile  -is [System.IO.FileInfo]) {$Referencefile  = $Referencefile.FullName}
    if ($Referencefile -and $Differencefile) {
         #if the filenames don't resolve, give up now.
         try     { $oneFile = ((Resolve-Path -Path $Referencefile -ErrorAction Stop).path -eq (Resolve-Path -Path $Differencefile  -ErrorAction Stop).path)}
         Catch   { Write-Warning -Message "Could not Resolve the filenames." ; return }

         #If we have one file , we must have two different Worksheet names. If we have two files $WorksheetName can be a single string or two strings.
         if      ($onefile -and ( ($WorksheetName.count -ne 2) -or $WorksheetName[0] -eq $WorksheetName[1] ) ) {
             Write-Warning -Message "If both the Reference and difference file are the same then Worksheet name must provide 2 different names"
             return
         }
         if      ($WorksheetName.count -eq 2)  {$Worksheet2 = $DiffPrefix = $WorksheetName[1] ; $Worksheet1 = $WorksheetName[0]  ;  }
         elseif  ($WorksheetName -is [string]) {$Worksheet2 = $Worksheet1 = $WorksheetName    ;
                                                $DiffPrefix = (Split-Path -Path $Differencefile -Leaf) -replace "\.xlsx$","" }
         else    {Write-Warning -Message "You must provide either a single Worksheet name or two names." ; return }

         $params= @{ ErrorAction = [System.Management.Automation.ActionPreference]::Stop }
         foreach ($p in @("HeaderName","NoHeader","StartRow")) {if ($PSBoundParameters[$p]) {$params[$p] = $PSBoundParameters[$p]}}
         try     {
             $ReferenceObject  = Import-Excel -Path $Referencefile  -WorksheetName $Worksheet1 @params
             $DifferenceObject = Import-Excel -Path $Differencefile -WorksheetName $Worksheet2 @Params
         }
         Catch   {Write-Warning -Message "Could not read the Worksheet from $Referencefile::$Worksheet1 and/or $Differencefile::$Worksheet2." ; return }
         if ($NoHeader) {$firstDataRow = $Startrow  } else {$firstDataRow = $Startrow + 1}
     }
     elseif (                $Differencefile) {
         if ($WorksheetName -isnot [string]) {Write-Warning -Message "You must provide a single Worksheet name." ; return }
         $params     =  @{WorksheetName=$WorksheetName; Path=$Differencefile; ErrorAction=[System.Management.Automation.ActionPreference]::Stop }
         foreach ($p in @("HeaderName","NoHeader","StartRow")) {if ($PSBoundParameters[$p]) {$params[$p] = $PSBoundParameters[$p]}}
         try            {$DifferenceObject = Import-Excel   @Params }
         Catch          {Write-Warning -Message "Could not read the Worksheet '$WorksheetName' from $Differencefile::$WorksheetName." ; return }
         if ($DiffPrefix -eq "=>" ) {
             $DiffPrefix  =  (Split-Path -Path $Differencefile -Leaf) -replace "\.xlsx$",""
         }
         if ($NoHeader) {$firstDataRow = $Startrow  } else {$firstDataRow = $Startrow + 1}
     }
     else   { $firstDataRow = 1  }
 #endregion

 #region Set lists of properties and row numbers
     #Make a list of properties/headings using the Property (default "*") and ExcludeProperty parameters
     $propList         = @()
     $DifferenceObject = $DifferenceObject | Update-FirstObjectProperties
     $headings         = $DifferenceObject[0].psobject.Properties.Name # This preserves the sequence - using get-member would sort them alphabetically! There may be extra properties in
     if ($NoHeader     -and "Name" -eq $Key)  {$Key     = "p1"}
     if ($headings     -notcontains    $Key -and
                              ('*' -ne $Key)) {Write-Warning -Message "You need to specify one of the headings in the sheet '$Worksheet1' as a key." ; return }
     foreach ($p in $Property)                { $propList += ($headings.where({$_ -like    $p}) )}
     foreach ($p in $ExcludeProperty)         { $propList  =  $propList.where({$_ -notlike $p})  }
     if (($propList    -notcontains $Key) -and
                           ('*' -ne $Key))    { $propList +=  $Key}    #If $key isn't one of the headings we will have bailed by now
     $propList         = $propList   | Select-Object -Unique           #so, prolist must contain at least $key if nothing else

     #If key is "*" we treat it differently , and we will create a script property which concatenates all the Properties in $Proplist
     $ConCatblock      = [scriptblock]::Create( ($proplist | ForEach-Object {'$this."' + $_ + '"'})  -join " + ")

     #Build the list of the properties to output, in order.
     $diffpart         = @()
     $refpart          = @()
     foreach ($p in $proplist.Where({$key -ne $_}) ) {$refPart += $p ; $diffPart += "$DiffPrefix $p" }
     $lastRefColNo     = $proplist.count
     $FirstDiffColNo   = $lastRefColNo + 1

     if ($key -ne '*') {
            $outputProps   = @($key) + $refpart + $diffpart
            #If we are using a single column as the key, don't duplicate it, so the last difference column will be A if there is one property, C if there are two, E if there are 3
            $lastDiffColNo = (2 * $proplist.count) - 1
     }
     else {
            $outputProps   = @( )    + $refpart + $diffpart
            #If we not using a single column as a key all columns are duplicated so, the Last difference column will be B if there is one property, D if there are two, F if there are 3
            $lastDiffColNo = (2 * $proplist.count )
     }

     #Add RowNumber to every row
     #If one sheet has extra rows we can get a single "==" result from compare, with the row from the reference sheet, but
     #the row in the other sheet might be different so we will look up the row number from the key field - build a hash table for that here
     #If we have "*" as the key ad the script property to concatenate the [selected] properties.

     $Rowhash = @{}
     $rowNo = $firstDataRow
     foreach ($row in $ReferenceObject)  {
        if   ($null -eq $row._row) {Add-Member -InputObject $row -MemberType NoteProperty   -Value ($rowNo ++)  -Name "_Row" }
        else {$rowNo++ }
        if   ($Key      -eq '*'  ) {Add-Member -InputObject $row -MemberType ScriptProperty -Value $ConCatblock -Name "_All" }
     }
     $rowNo = $firstDataRow
     foreach ($row in $DifferenceObject) {
         Add-Member       -InputObject $row -MemberType NoteProperty   -Value $rowNo       -Name "$DiffPrefix Row" -Force
         if   ($Key       -eq '*' )    {
               Add-Member -InputObject $row -MemberType ScriptProperty -Value $ConCatblock -Name "_All"
               $Rowhash[$row._All] = $rowNo
         }
         else {$Rowhash[$row.$key] = $rowNo  }
         $rowNo ++
     }
     if ($DifferenceObject.count -gt $Rowhash.Keys.Count) {
        Write-Warning -Message "Difference object has $($DifferenceObject.Count) rows; but only $($Rowhash.keys.count) unique keys"
     }
     if ($Key -eq '*') {$key = "_ALL"}
 #endregion
     #We need to know all the properties we've met on the objects we've diffed
     $eDiffProps  = [ordered]@{}
     #When we do a compare object changes will result in two rows so we group them and join them together.
     $expandedDiff = Compare-Object -ReferenceObject $ReferenceObject -DifferenceObject $DifferenceObject -Property $propList -PassThru -IncludeEqual |
                        Group-Object -Property $key | ForEach-Object {
                            #The value of the key column is the name of the Group.
                            $keyval = $_.name
                            #we're going to create a custom object from a hash table.
                            $hash = [ordered]@{}
                            foreach ($result in $_.Group) {
                                if     ($result.SideIndicator -ne "=>")      {$hash["_Row"] = $result._Row  }
                                elseif (-not $hash["$DiffPrefix Row"])       {$hash["_Row"] = "" }
                                #if we have already set the side, this must be the second record, so set side to indicate "changed"; if we got two "Same" indicators we may have a classh of keys
                                if     ($hash.Side) {
                                    if ($hash.Side -eq $result.SideIndicator) {Write-Warning -Message "'$keyval' may be a duplicate."}
                                        $hash.Side = "<>"
                                }
                                else   {$hash["Side"] = $result.SideIndicator}
                                switch ($hash.side) {
                                    '==' {      $hash["$DiffPrefix is"] = 'Same'   }
                                    '=>' {      $hash["$DiffPrefix is"] = 'Added'  }
                                    '<>' { if (-not $hash["_Row"]) {
                                                $hash["$DiffPrefix is"] = 'Added'
                                            }
                                            else {
                                                $hash["$DiffPrefix is"] = 'Changed'
                                            }
                                         }
                                    '<=' {      $hash["$DiffPrefix is"] = 'Removed'}
                                    }
                                 #find the number of the row in the the "difference" object which has this key. If it is the object is only in the reference this will be blank.
                                 $hash["$DiffPrefix Row"] = $Rowhash[$keyval]
                                 $hash[$key]              = $keyval
                                 #Create FieldName and/or =>FieldName columns
                                 foreach  ($p in $result.psobject.Properties.name.where({$_ -ne $key -and $_ -ne "SideIndicator" -and $_ -ne "$DiffPrefix Row" })) {
                                    if     ($result.SideIndicator -eq "==" -and $p -in $propList)
                                                                             {$hash[("$p")] = $hash[("$DiffPrefix $p")] = $result.$P}
                                    elseif ($result.SideIndicator -eq "==" -or $result.SideIndicator -eq "<=")
                                                                             {$hash[("$p")]                             = $result.$P}
                                    elseif ($result.SideIndicator -eq "=>")  {                $hash[("$DiffPrefix $p")] = $result.$P}
                                 }
                             }

                             foreach ($k in $hash.keys) {$eDiffProps[$k] = $true}
                             [Pscustomobject]$hash
     }  | Sort-Object -Property "_row"

     #Already sorted by reference row number, fill in any blanks in the difference-row column.
     for ($i = 1; $i -lt $expandedDiff.Count; $i++) {if (-not $expandedDiff[$i]."$DiffPrefix Row") {$expandedDiff[$i]."$DiffPrefix Row" = $expandedDiff[$i-1]."$DiffPrefix Row" } }

     #Now re-Sort by difference row number, and fill in any blanks in the reference-row column.
     $expandedDiff = $expandedDiff | Sort-Object -Property "$DiffPrefix Row"
     for ($i = 1; $i -lt $expandedDiff.Count; $i++) {if (-not $expandedDiff[$i]."_Row") {$expandedDiff[$i]."_Row" = $expandedDiff[$i-1]."_Row" } }

     $AllProps = @("_Row") + $OutputProps + $eDiffProps.keys.where({$_ -notin ($outputProps + @("_row","side","SideIndicator","_ALL" ))})

     if     ($PassThru -or -not $OutputFile) {return  ($expandedDiff | Select-Object -Property $allprops  | Sort-Object -Property  "_row", "$DiffPrefix Row"    )  }
     elseif ($PSCmdlet.ShouldProcess($OutputFile,"Write Output to Excel file")) {
         $expandedDiff =  $expandedDiff | Sort-Object -Property  "_row", "$DiffPrefix Row"
         $xl = $expandedDiff | Select-Object -Property   $OutputProps    | Update-FirstObjectProperties      |
           Export-Excel -Path $OutputFile -Worksheetname $OutputSheetName -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -PassThru
         $ws =  $xl.Workbook.Worksheets[$OutputSheetName]
         for ($i = 0; $i -lt $expandedDiff.Count; $i++ ) {
            if     ( $expandedDiff[$i].side -ne "==" )  {
                Set-ExcelRange -Worksheet $ws     -Range ("A" + ($i + 2 )) -FontColor       $KeyFontColor
            }
            elseif ( $HideEqual                      )  {$ws.row($i+2).hidden = $true }
            if     ( $expandedDiff[$i].side -eq "<>" )  {
                $range = $ws.Dimension -replace "\d+",  ($i + 2 )
                Set-ExcelRange -Worksheet $ws     -Range $range            -BackgroundColor $ChangeBackgroundColor
            }
            elseif ( $expandedDiff[$i].side -eq "<=" )  {
                $rangeR1C1 = "R[{0}]C[1]:R[{0}]C[{1}]" -f ($i + 2 ) , $lastRefColNo
                $range = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1($rangeR1C1,0,0)
                Set-ExcelRange -Worksheet $ws     -Range $range            -BackgroundColor $DeleteBackgroundColor
            }
            elseif ( $expandedDiff[$i].side -eq "=>" )  {
                if ($propList.count -gt 1) {
                    $rangeR1C1 = "R[{0}]C[{1}]:R[{0}]C[{2}]" -f ($i + 2 ) , $FirstDiffColNo , $lastDiffColNo
                    $range = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1($rangeR1C1,0,0)
                    Set-ExcelRange -Worksheet $ws -Range $range            -BackgroundColor $AddBackgroundColor
                }
                Set-ExcelRange -Worksheet $ws     -Range ("A" + ($i + 2 )) -BackgroundColor $AddBackgroundColor
            }
         }
         Close-ExcelPackage -ExcelPackage $xl -Show:$Show
     }
}

Function Merge-MultipleSheets {
   <#
      .Synopsis
        Merges Worksheets into a single Worksheet with differences marked up.
      .Description
        The Merge Worksheet command combines two sheets. Merge-MultipleSheets is
        designed to merge more than two. So if asked to merge sheets A,B,C  which
        contain Services, with a Name, Displayname and Start mode, where "Name" is
        treated as the key, Merge-MultipleSheets calls Merge-Worksheet to merge
        "Name", "Displayname" and "Startmode" from sheets A and C;  the result has
        column headings  "_Row", "Name", "DisplayName", "Startmode", "C-DisplayName",
        "C-StartMode", "C-Is" and "C-Row".
        Merge-MultipleSheets then calls Merge-Worksheet again passing it the
        intermediate result and sheet B, comparing "Name", "Displayname" and
        "Start mode" columns on each side, and gets a result with columns "_Row",
        "Name", "DisplayName", "Startmode", "B-DisplayName",  "B-StartMode", "B-Is",
        "B-Row", "C-DisplayName", "C-StartMode", "C-Is" and "C-Row". Any columns on
        the "reference" side which are not used in the comparison are added on the
        right, which is why we compare the sheets in reverse order.

        The "Is" columns hold "Same", "Added", "Removed" or "Changed" and is used for
        conditional formatting in the output sheet (these columns are hidden by default),
        and when the data is written to Excel the "reference" columns, in this case
        "DisplayName" and "Start" are renamed to reflect their source, so become
        "A-DisplayName" and "A-Start".

        Conditional formatting is also applied to the Key column ("Name" in this
        case) so the view can be filtered to rows with changes by filtering this
        column on color.

        Note: the processing order can affect what is seen as a change. For example
        if there is an extra item in sheet B in the example above, Sheet C will be
        processed and that row and will not be seen to be missing. When sheet B is
        processed it is marked as an addition, and the conditional formatting marks
        the entries from sheet A to show that a values were added in at least one
        sheet.   However if Sheet B is the reference sheet, A and C will be seen to
        have an item removed; and if B is processed before C, the extra item is
        known when C is processed and so C is considered to be missing that item.
      .Example
        dir Server*.xlsx | Merge-MulipleSheets   -WorksheetName Services -OutputFile Test2.xlsx -OutputSheetName Services -Show
        Here we are auditing servers and each one has a workbook in the current
        directory which contains a "Services" Worksheet (the result of
        Get-WmiObject -Class win32_service  | Select-Object -Property Name, Displayname, Startmode)
        No key is specified so the key is assumed to be the "Name" column.
        The files are merged and the result is opened on completion.
      .Example
        dir Serv*.xlsx |  Merge-MulipleSheets  -WorksheetName Software -Key "*" -ExcludeProperty Install* -OutputFile Test2.xlsx -OutputSheetName Software -Show
        The server audit files in the previous example also have "Software" worksheet,
        but no single field on that sheet works as a key. Specifying "*" for the key
        produces a compound key using all non-excluded fields (and the installation
        date and file location are excluded).
      .Example
        Merge-MulipleSheets -Path hotfixes.xlsx -WorksheetName Serv* -Key hotfixid -OutputFile test2.xlsx -OutputSheetName hotfixes  -HideRowNumbers -Show
        This time all the servers have written their hotfix information to their own
        worksheets in a shared Excel workbook named "Hotfixes.xlsx" (the information was
        obtained by running Get-Hotfix | Sort-Object -Property description,hotfixid  | Select-Object -Property Description,HotfixID)
        This ignores any sheets which are not named "Serv*", and uses the HotfixID as
        the key; in this version the row numbers are hidden.
    #>
    [cmdletbinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification="False positives when initializing variable in begin block")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification="MultipleSheet would be incorrect")]
    #[Alias("Merge-MulipleSheets")] #There was a spelling error in the first release. This was there to ensure things didn't break but intelisense gave the alias first.
    param   (
        #Paths to the files to be merged. Files are also accepted
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $Path  ,
        #The row from where we start to import data, all rows above the Start row are disregarded. By default this is the first row.
        [int]$Startrow = 1,

        #Specifies custom property names to use, instead of the values defined in the column headers of the Start row.
        [String[]]$Headername,

        #If specified, property names will be automatically generated (P1, P2, P3, ..) instead of using the values from the start row.
        [switch]$NoHeader,

        #Name(s) of Worksheets to compare.
        $WorksheetName   = "Sheet1",
        #File to write output to.
        [Alias('OutFile')]
        $OutputFile = ".\temp.xlsx",
        #Name of Worksheet to output - if none specified will use the reference Worksheet name.
        [Alias('OutSheet')]
        $OutputSheetName = "Sheet1",
        #Properties to include in the comparison - supports wildcards, default is "*".
        $Property        = "*"    ,
        #Properties to exclude from the the comparison - supports wildcards.
        $ExcludeProperty ,
        #Name of a column which is unique used to pair up rows from the reference and difference sides, default is "Name".
        $Key           = "Name" ,
        #Sets the font color for the Key field; this means you can filter by color to get only changed rows.
        $KeyFontColor          = [System.Drawing.Color]::Red,
        #Sets the background color for changed rows.
        $ChangeBackgroundColor = [System.Drawing.Color]::Orange,
        #Sets the background color for rows in the reference but deleted from the difference sheet.
        $DeleteBackgroundColor = [System.Drawing.Color]::LightPink,
        #Sets the background color for rows not in the reference but added to the difference sheet.
        $AddBackgroundColor    = [System.Drawing.Color]::Orange,
        #If specified, hides the columns in the spreadsheet that contain the row numbers.
        [switch]$HideRowNumbers ,
        #If specified, outputs the data to the pipeline (you can add -whatif so it the command only outputs to the pipeline).
        [switch]$Passthru  ,
        #If specified, opens the output workbook.
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
                                  Export-Excel -Path $OutputFile -Worksheetname $OutputSheetName -ClearSheet -BoldTopRow -AutoFilter -PassThru
        $sheet                 = $excel.Workbook.Worksheets[$OutputSheetName]

        #We will put in a conditional format for "if all the others are not flagged as 'same'" to mark rows where something is added, removed or changed
        $sameChecks            = @()

        #All the 'difference' columns in the sheet are labeled with the file they came from, 'reference' columns need their
        #headers prefixed with the ref file name,  $colnames is the basis of a regular expression to identify what should have $refPrefix appended
        $colNames              = @("^_Row$")
        if ($key -ne "*")
              {$colnames      += "^$Key$"}
        if ($filesToProcess.Count -ge 2) {
              $refPrefix       = (Split-Path -Path $filestoProcess[0] -Leaf) -replace "\.xlsx$"," "
        }
        else {$refPrefix       = $WorksheetName[0] }
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
        $sheet.Cells.AutoFitColumns()

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
