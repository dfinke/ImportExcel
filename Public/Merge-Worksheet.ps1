function Merge-Worksheet {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
         [parameter(ParameterSetName='A',Mandatory=$true,Position=0)]  #A = Compare two files default headers
         [parameter(ParameterSetName='B',Mandatory=$true,Position=0)]  #B = Compare two files user supplied headers
         [parameter(ParameterSetName='C',Mandatory=$true,Position=0)]  #C = Compare two files headers P1, P2, P3 etc
         $Referencefile ,

         [parameter(ParameterSetName='A',Mandatory=$true,Position=1)]
         [parameter(ParameterSetName='B',Mandatory=$true,Position=1)]
         [parameter(ParameterSetName='C',Mandatory=$true,Position=1)]
         [parameter(ParameterSetName='E',Mandatory=$true,Position=1)] #D Compare two objects; E = Compare one object one file that uses default headers
         [parameter(ParameterSetName='F',Mandatory=$true,Position=1)] #F = Compare one object one file that uses user supplied headers
         [parameter(ParameterSetName='G',Mandatory=$true,Position=1)] #G   Compare one object one file that uses headers P1, P2, P3 etc
         $Differencefile ,

         [parameter(ParameterSetName='A',Position=2)]  #Applies to all sets EXCEPT D which is two objects (no sheets)
         [parameter(ParameterSetName='B',Position=2)]
         [parameter(ParameterSetName='C',Position=2)]
         [parameter(ParameterSetName='E',Position=2)]
         [parameter(ParameterSetName='F',Position=2)]
         [parameter(ParameterSetName='G',Position=2)]
         $WorksheetName   = "Sheet1",

         [parameter(ParameterSetName='A')]  #Applies to all sets EXCEPT D which is two objects (no sheets, so no start row )
         [parameter(ParameterSetName='B')]
         [parameter(ParameterSetName='C')]
         [parameter(ParameterSetName='E')]
         [parameter(ParameterSetName='F')]
         [parameter(ParameterSetName='G')]
         [int]$Startrow = 1,

         [Parameter(ParameterSetName='B',Mandatory=$true)]  #Compare  object + sheet or 2 sheets with user supplied headers
         [Parameter(ParameterSetName='F',Mandatory=$true)]
         [String[]]$Headername,

         [Parameter(ParameterSetName='C',Mandatory=$true)]  #Compare  object + sheet or 2 sheets with headers of P1, P2, P3 ...
         [Parameter(ParameterSetName='G',Mandatory=$true)]
         [switch]$NoHeader,

         [parameter(ParameterSetName='D',Mandatory=$true)]
         [parameter(ParameterSetName='E',Mandatory=$true)]
         [parameter(ParameterSetName='F',Mandatory=$true)]
         [parameter(ParameterSetName='G',Mandatory=$true)]
         [Alias('RefObject')]
         $ReferenceObject ,
         [parameter(ParameterSetName='D',Mandatory=$true,Position=1)]
         [Alias('DiffObject')]
         $DifferenceObject ,
         [parameter(ParameterSetName='D',Position=2)]
         [parameter(ParameterSetName='E',Position=2)]
         [parameter(ParameterSetName='F',Position=2)]
         [parameter(ParameterSetName='G',Position=2)]
         $DiffPrefix = "=>" ,
         [parameter(Position=3)]
         [Alias('OutFile')]
         $OutputFile ,
         [parameter(Position=4)]
         [Alias('OutSheet')]
         $OutputSheetName = "Sheet1",
         $Property        = "*"    ,
         $ExcludeProperty ,
         $Key           = "Name" ,
         $KeyFontColor          = [System.Drawing.Color]::DarkRed ,
         $ChangeBackgroundColor = [System.Drawing.Color]::Orange,
         $DeleteBackgroundColor = [System.Drawing.Color]::LightPink,
         $AddBackgroundColor    = [System.Drawing.Color]::PaleGreen,
         [switch]$HideEqual ,
         [switch]$Passthru  ,
         [Switch]$Show
    )

 #region Read Excel data
    if ($Differencefile -is [System.IO.FileInfo]) {$Differencefile = $Differencefile.FullName}
    if ($Referencefile  -is [System.IO.FileInfo]) {$Referencefile  = $Referencefile.FullName}
    if ($Referencefile -and $Differencefile) {
         #if the filenames don't resolve, give up now.
         try     { $oneFile = ((Resolve-Path -Path $Referencefile -ErrorAction Stop).path -eq (Resolve-Path -Path $Differencefile  -ErrorAction Stop).path)}
         catch   { Write-Warning -Message "Could not Resolve the filenames." ; return }

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
         catch   {Write-Warning -Message "Could not read the Worksheet from $Referencefile::$Worksheet1 and/or $Differencefile::$Worksheet2." ; return }
         if ($NoHeader) {$firstDataRow = $Startrow  } else {$firstDataRow = $Startrow + 1}
     }
     elseif (                $Differencefile) {
         if ($WorksheetName -isnot [string]) {Write-Warning -Message "You must provide a single Worksheet name." ; return }
         $params     =  @{WorksheetName=$WorksheetName; Path=$Differencefile; ErrorAction=[System.Management.Automation.ActionPreference]::Stop }
         foreach ($p in @("HeaderName","NoHeader","StartRow")) {if ($PSBoundParameters[$p]) {$params[$p] = $PSBoundParameters[$p]}}
         try            {$DifferenceObject = Import-Excel   @Params }
         catch          {Write-Warning -Message "Could not read the Worksheet '$WorksheetName' from $Differencefile::$WorksheetName." ; return }
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
     $headings         = $DifferenceObject[0].psobject.Properties.Name # This preserves the sequence - using Get-member would sort them alphabetically! There may be extra properties in
     if ($NoHeader     -and "Name" -eq $Key)  {$Key     = "p1"}
     if ($headings     -notcontains    $Key -and
                              ('*' -ne $Key)) {Write-Warning -Message "You need to specify one of the headings in the sheet '$Worksheet2' as a key." ; return }
     foreach ($p in $Property)                { $propList += ($headings.where({$_ -like    $p}) )}
     foreach ($p in $ExcludeProperty)         { $propList  =  $propList.where({$_ -notlike $p})  }
     if (($propList    -notcontains $Key) -and
                           ('*' -ne $Key))    { $propList +=  $Key}    #If $Key isn't one of the headings we will have bailed by now
     $propList         = $propList   | Select-Object -Unique           #so, prolist must contain at least $Key if nothing else

     #If key is "*" we treat it differently , and we will create a script property which concatenates all the Properties in $Proplist
     $ConCatblock      = [scriptblock]::Create( ($proplist | ForEach-Object {'$this."' + $_ + '"'})  -join " + ")

     #Build the list of the properties to output, in order.
     $diffpart         = @()
     $refpart          = @()
     foreach ($p in $proplist.Where({$Key -ne $_}) ) {$refPart += $p ; $diffPart += "$DiffPrefix $p" }
     $lastRefColNo     = $proplist.count
     $FirstDiffColNo   = $lastRefColNo + 1

     if ($Key -ne '*') {
            $outputProps   = @($Key) + $refpart + $diffpart
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

     $rowHash = @{}
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
               $rowHash[$row._All] = $rowNo
         }
         else {$rowHash[$row.$Key] = $rowNo  }
         $rowNo ++
     }
     if ($DifferenceObject.count -gt $rowHash.Keys.Count) {
        Write-Warning -Message "Difference object has $($DifferenceObject.Count) rows; but only $($rowHash.keys.count) unique keys"
     }
     if ($Key -eq '*') {$Key = "_ALL"}
 #endregion
     #We need to know all the properties we've met on the objects we've diffed
     $eDiffProps  = [ordered]@{}
     #When we do a compare object changes will result in two rows so we group them and join them together.
     $expandedDiff = Compare-Object -ReferenceObject $ReferenceObject -DifferenceObject $DifferenceObject -Property $propList -PassThru -IncludeEqual |
                        Group-Object -Property $Key | ForEach-Object {
                            #The value of the key column is the name of the Group.
                            $keyVal = $_.name
                            #we're going to create a custom object from a hash table.
                            $hash = [ordered]@{}
                            foreach ($result in $_.Group) {
                                if     ($result.SideIndicator -ne "=>")      {$hash["_Row"] = $result._Row  }
                                elseif (-not $hash["$DiffPrefix Row"])       {$hash["_Row"] = "" }
                                #if we have already set the side, this must be the second record, so set side to indicate "changed"; if we got two "Same" indicators we may have a classh of keys
                                if     ($hash.Side) {
                                    if ($hash.Side -eq $result.SideIndicator) {Write-Warning -Message "'$keyVal' may be a duplicate."}
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
                                 $hash["$DiffPrefix Row"] = $rowHash[$keyVal]
                                 $hash[$Key]              = $keyVal
                                 #Create FieldName and/or =>FieldName columns
                                 foreach  ($p in $result.psobject.Properties.name.where({$_ -ne $Key -and $_ -ne "SideIndicator" -and $_ -ne "$DiffPrefix Row" })) {
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
           Export-Excel -Path $OutputFile -WorksheetName $OutputSheetName -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -PassThru
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
