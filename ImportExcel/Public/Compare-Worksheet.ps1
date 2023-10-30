function Compare-Worksheet {
    [CmdletBinding(DefaultParameterSetName)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification="Write host used for sub-warning level message to operator which does not form output")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification="False positives when initializing variable in begin block")]
    param(
        [parameter(Mandatory=$true,Position=0)]
        $Referencefile ,
        [parameter(Mandatory=$true,Position=1)]
        $Differencefile   ,
        $WorksheetName   = "Sheet1",
        $Property        = "*"    ,
        $ExcludeProperty ,
        [Parameter(ParameterSetName='B', Mandatory)]
        [String[]]$Headername,
        [Parameter(ParameterSetName='C', Mandatory)]
        [switch]$NoHeader,
        [int]$Startrow = 1,
        $AllDataBackgroundColor,
        $BackgroundColor,
        $TabColor,
        $Key             = "Name" ,
        $FontColor,
        [Switch]$Show,
        [switch]$GridView,
        [Switch]$PassThru,
        [Switch]$IncludeEqual,
        [Switch]$ExcludeDifferent
    )

    #if the filenames don't resolve, give up now.
    try    { $oneFile = ((Resolve-Path -Path $Referencefile -ErrorAction Stop).path -eq (Resolve-Path -Path $Differencefile  -ErrorAction Stop).path)}
    catch  { Write-Warning -Message "Could not Resolve the filenames." ; return }

    #If we have one file , we must have two different worksheet names. If we have two files we can have a single string or two strings.
    if     ($onefile -and ( ($WorksheetName.count -ne 2) -or $WorksheetName[0] -eq $WorksheetName[1] ) ) {
        Write-Warning -Message "If both the Reference and difference file are the same then worksheet name must provide 2 different names"
        return
    }
    if     ($WorksheetName.count -eq 2)       {$worksheet1 = $WorksheetName[0] ;   $worksheet2 = $WorksheetName[1]}
    elseif ($WorksheetName -is [string])      {$worksheet1 = $worksheet2 = $WorksheetName}
    else   {Write-Warning -Message "You must provide either a single worksheet name or two names." ; return }

    $params= @{ ErrorAction = [System.Management.Automation.ActionPreference]::Stop }
    foreach ($p in @("HeaderName","NoHeader","StartRow")) {if ($PSBoundParameters[$p]) {$params[$p] = $PSBoundParameters[$p]}}
    try    {
        $sheet1 = Import-Excel -Path $Referencefile  -WorksheetName $worksheet1 @params
        $sheet2 = Import-Excel -Path $Differencefile -WorksheetName $worksheet2 @Params
    }
    catch  {Write-Warning -Message "Could not read the worksheet from $Referencefile and/or $Differencefile." ; return }

    #Get Column headings and create a hash table of Name to column letter.
    $headings = $Sheet1[-1].psobject.Properties.name # This preserves the sequence - using Get-member would sort them alphabetically!
    $headings | ForEach-Object -Begin {$columns  = @{}  ; $i= 1 } -Process  {$Columns[$_] = [OfficeOpenXml.ExcelAddress]::GetAddress(1,($i ++)) -replace "\d","" }

    #Make a list of property headings using the Property (default "*") and ExcludeProperty parameters
    if ($Key -eq "Name" -and $NoHeader) {$Key  = "p1"}
    $propList = @()
    foreach ($p in $Property)           {$propList += ($headings.where({$_ -like    $p}) )}
    foreach ($p in $ExcludeProperty)    {$propList  =  $propList.where({$_ -notlike $p})  }
    if (($headings -contains $Key) -and ($propList -notcontains $Key)) {$propList += $Key}
    $propList = $propList | Select-Object -Unique
    if ($propList.Count -eq 0)  {Write-Warning -Message "No Columns are selected with -Property = '$Property' and -excludeProperty = '$ExcludeProperty'." ; return}

    #Add RowNumber, Sheetname and file name to every row
    $firstDataRow = $startRow + 1
    if ($Headername -or $NoHeader) {$firstDataRow -- }
    $i = $firstDataRow ; foreach ($row in $Sheet1) {Add-Member -InputObject $row -MemberType NoteProperty -Name "_Row"   -Value ($i ++)
                                                    Add-Member -InputObject $row -MemberType NoteProperty -Name "_Sheet" -Value  $worksheet1
                                                    Add-Member -InputObject $row -MemberType NoteProperty -Name "_File"  -Value  $Referencefile}
    $i = $firstDataRow ; foreach ($row in $Sheet2) {Add-Member -InputObject $row -MemberType NoteProperty -Name "_Row"   -Value ($i ++)
                                                    Add-Member -InputObject $row -MemberType NoteProperty -Name "_Sheet" -Value  $worksheet2
                                                    Add-Member -InputObject $row -MemberType NoteProperty -Name "_File"  -Value  $Differencefile}

    if ($ExcludeDifferent -and -not $IncludeEqual) {$IncludeEqual = $true}
    #Do the comparison and add file,sheet and row to the result - these are prefixed with "_" to show they are added the addition will fail if the sheet has these properties so split the operations
    [PSCustomObject[]]$diff = Compare-Object -ReferenceObject $Sheet1 -DifferenceObject $Sheet2 -Property $propList -PassThru -IncludeEqual:$IncludeEqual -ExcludeDifferent:$ExcludeDifferent  |
                Sort-Object -Property "_Row","File"

    #if BackgroundColor was specified, set it on extra or extra or changed rows
    if      ($diff -and $BackgroundColor) {
        #Differences may only exist in one file. So gather the changes for each file; open the file, update each impacted row in the shee, save the file
        $updates = $diff.where({$_.SideIndicator -ne "=="}) | Group-object -Property "_File"
        foreach   ($file in $updates) {
            try   {$xl  = Open-ExcelPackage -Path $file.name }
            catch {Write-warning -Message "Can't open $($file.Name) for writing." ; return}
            if  ($PSBoundParameters.ContainsKey("AllDataBackgroundColor")) {
                $file.Group._sheet | Sort-Object -Unique | ForEach-Object {
                    $ws =  $xl.Workbook.Worksheets[$_]
                    if ($headerName) {$range = "A" +  $startrow      + ":" + $ws.dimension.end.address}
                    else             {$range = "A" + ($startrow + 1) + ":" + $ws.dimension.end.address}
                    Set-ExcelRange -Worksheet $ws -BackgroundColor $AllDataBackgroundColor -Range $Range
                }
            }
            foreach ($row in $file.group)  {
                $ws    = $xl.Workbook.Worksheets[$row._Sheet]
                $range = $ws.Dimension -replace "\d+",$row._row
                Set-ExcelRange -Worksheet $ws -Range $range -BackgroundColor $BackgroundColor
            }
            if  ($PSBoundParameters.ContainsKey("TabColor")) {
                if ($TabColor -is [string])         {$TabColor = [System.Drawing.Color]::$TabColor }
                foreach ($tab in ($file.group._sheet | Select-Object -Unique)) {
                    $xl.Workbook.Worksheets[$tab].TabColor = $TabColor
                 }
            }
            $xl.save()  ; $xl.Stream.Close() ; $xl.Dispose()
        }
    }
    #if font color was specified, set it on changed properties where the same key appears in both sheets.
    if      ($diff -and $FontColor -and (($propList -contains $Key) -or ($Key -is [hashtable]))  ) {
        $updates = $diff.where({$_.SideIndicator -ne "=="})  | Group-object -Property $Key | Where-Object {$_.count -eq 2}
        if ($updates) {
            $XL1 = Open-ExcelPackage -path $Referencefile
            if ($oneFile ) {$xl2 = $xl1}
            else           {$xl2 = Open-ExcelPackage -path $Differencefile }
            foreach ($u in $updates) {
                 foreach ($p in $propList) {
                    if ($u.group[0]._file -eq $Referencefile) {
                        $ws1 =  $xl1.Workbook.Worksheets[$u.Group[0]._sheet]
                        $ws2 =  $xl2.Workbook.Worksheets[$u.Group[1]._sheet]
                    }
                    else {
                        $ws1 =  $xl2.Workbook.Worksheets[$u.Group[0]._sheet]
                        $ws2 =  $xl1.Workbook.Worksheets[$u.Group[1]._sheet]
                    }
                    if($u.Group[0].$p -ne $u.Group[1].$p ) {
                        Set-ExcelRange -Worksheet $ws1 -Range ($Columns[$p] + $u.Group[0]._Row) -FontColor $FontColor
                        Set-ExcelRange -Worksheet $ws2 -Range ($Columns[$p] + $u.Group[1]._Row) -FontColor $FontColor
                    }
                }
            }
            $xl1.Save()                     ; $xl1.Stream.Close() ; $xl1.Dispose()
            if (-not $oneFile) {$xl2.Save() ; $xl2.Stream.Close() ; $xl2.Dispose()}
        }
    }
    elseif  ($diff -and $FontColor) {Write-Warning -Message "To match rows to set changed cells, you must specify -Key and it must match one of the included properties." }

    #if nothing was found write a message which will not be redirected
    if (-not $diff) {Write-Host "Comparison of $Referencefile::$worksheet1 and $Differencefile::$worksheet2 returned no results."  }

    if      ($Show)               {
        Start-Process -FilePath $Referencefile
        if  (-not $oneFile)  { Start-Process -FilePath $Differencefile }
        if  ($GridView)      { Write-Warning -Message "-GridView is ignored when -Show is specified" }
    }
    elseif  ($GridView -and $propList -contains $Key) {


             if ($IncludeEqual -and -not $ExcludeDifferent) {
                $GroupedRows = $diff | Group-Object -Property $Key
             }
             else { #to get the right now numbers on the grid we need to have all the rows.
                $GroupedRows = Compare-Object -ReferenceObject $Sheet1 -DifferenceObject $Sheet2 -Property $propList -PassThru -IncludeEqual  |
                                        Group-Object -Property $Key
             }
             #Additions, deletions and unchanged rows will give a group of 1; changes will give a group of 2 .

             #If one sheet has extra rows we can get a single "==" result from compare, but with the row from the reference sheet
             #but the row in the other sheet might so we will look up the row number from the key field build a hash table for that
             $Sheet2 | ForEach-Object -Begin {$rowHash = @{} } -Process {$rowHash[$_.$Key] = $_._row }

             $ExpandedDiff = ForEach ($g in $GroupedRows)  {
                #we're going to create a custom object from a hash table. We want the fields to be ordered
                $hash = [ordered]@{}
                foreach ($result IN $g.Group) {
                    # if result indicates equal or "in Reference" set the reference side row. If we did that on a previous result keep it. Otherwise set to "blank"
                    if     ($result.sideindicator -ne "=>")      {$hash["<Row"] = $result._Row  }
                    elseif (-not $hash["<Row"])                  {$hash["<Row"] = "" }
                    #if we have already set the side, this is the second record, so set side to indicate "changed"
                    if     ($hash.Side) {$hash.side = "<>"} else {$hash["Side"] = $result.sideindicator}
                    #if result is "in reference" and we don't have a matching "in difference" (meaning a change) the lookup will be blank. Which we want.
                    $hash[">Row"] = $rowHash[$g.Name]
                    #position the key as the next field (only appears once)
                    $Hash[$Key]    = $g.Name
                    #For all the other fields we care about create <=FieldName and/or =>FieldName
                    foreach ($p in $propList.Where({$_ -ne $Key})) {
                        if  ($result.SideIndicator -eq "==")  {$hash[("=>$P")] = $hash[("<=$P")] =$result.$P}
                        else                                  {$hash[($result.SideIndicator+$P)] =$result.$P}
                    }
                }
                [Pscustomobject]$hash
             }

             #Sort by reference row number, and fill in any blanks in the difference-row column
             $ExpandedDiff = $ExpandedDiff | Sort-Object -Property "<row"
             for ($i = 1; $i -lt $ExpandedDiff.Count; $i++) {if (-not $ExpandedDiff[$i].">row") {$ExpandedDiff[$i].">row" = $ExpandedDiff[$i-1].">row" } }
             #Sort by difference row number, and fill in any blanks in the reference-row column
             $ExpandedDiff = $ExpandedDiff | Sort-Object -Property ">row"
             for ($i = 1; $i -lt $ExpandedDiff.Count; $i++) {if (-not $ExpandedDiff[$i]."<row") {$ExpandedDiff[$i]."<row" = $ExpandedDiff[$i-1]."<row" } }

             #if we had to put the equal rows back, take them out; sort, make sure all the columns are present in row 1 so the grid puts them in, and output
             if ( $ExcludeDifferent) {$ExpandedDiff = $ExpandedDiff.where({$_.side -eq "=="}) | Sort-Object -Property "<row" ,">row"  }
             elseif ( $IncludeEqual) {$ExpandedDiff = $ExpandedDiff                           | Sort-Object -Property "<row" ,">row"  }
             else                    {$ExpandedDiff = $ExpandedDiff.where({$_.side -ne "=="}) | Sort-Object -Property "<row" ,">row"  }
             $ExpandedDiff | Update-FirstObjectProperties | Out-GridView -Title "Comparing $Referencefile::$worksheet1 (<=) with $Differencefile::$worksheet2 (=>)"
    }
    elseif  ($GridView     )  {Write-Warning -Message "To use -GridView you must specify -Key and it must match one of the included properties."  }
    elseif  (-not $PassThru)  {return ($diff | Select-Object -Property (@(@{n="_Side";e={$_.SideIndicator}},"_File" ,"_Sheet","_Row") + $propList))}
    if      (     $PassThru)  {return  $diff }
}
