Function Compare-Worksheet { 
<#
    .Synopsis 
        Compares two worksheets with the same name in different files. 
    .Description
        This command takes two file names, a worksheet name and a name for a key column. 
        It reads the worksheet from each file and decides the column names.
        It builds as hashtable of the key column values and the rows they appear in  
        It then uses PowerShell's compare object command to compare the sheets (explicity checking all column names which have not been excluded)
        For the difference rows it adds the row number for the key of that row - we have to add the key after doing the comparison, 
        otherwise rows will be considered as different simply because they have different row numbers 
        We also add the name of the file in which the difference occurs.  
        If -BackgroundColor is specified the difference rows will be changed to that background. 
    .Example 
        Compare-WorkSheet -Referencefile 'Server1.xlsx' -Differencefile 'Server2.xlsx'  -WorkSheetName Products -key IdentifyingNumber -ExcludeProperty Install* | format-table
        The two workbooks in this example contain the result of redirecting a subset of properties from Get-WmiObject -Class win32_product to Export-Excel
        The command compares the "products" pages in the two workbooks, but we don't want a match if the software was installed on a 
        different date or from a different place,  so Excluding Install* removes InstallDate and InstallSource. The results will be presented as a table.  
    .Example 
        Compare-WorkSheet  'Server1.xlsx' 'Server2.xlsx'  -WorkSheetName Services -key Name -BackgroundColor lightGreen
        This time two workbooks contain the result of redirecting Get-WmiObject -Class win32_service to Export-Excel 
        This command compares the "services" pages and highlights the rows in the spreadsheet files. 
        Here the -Differencefile and -Referencefile parameter switches are assumed
    .Example 
        Compare-WorkSheet 'Server1.xlsx' 'Server2.xlsx'  -WorkSheetName Services -BackgroundColor lightGreen -fontColor Red -Show
        This builds on the previous example: this time Where two rows in the services have the same name, this will also highlight  the changed cells in red. 
        This example will open the Excel files and  omits the -key parameter because "Name" will be assumed to the label for the key column 
    .Example
        Compare-WorkSheet 'Pester-tests.xlsx' 'Pester-tests.xlsx' -WorkSheetName 'Server1','Server2' -Property "full Description","Executed","Result" -Key "full Description" -FontColor Red -TabColor Yellow -Show
        This time the reference file and the difference file are the same file and two different sheets are used. Because the tests include the
        machine name and time the test was run only a limited set of columns.   
    .Example
        Compare-WorkSheet - 'Server1.xlsx' 'Server2.xlsx' -WorkSheetName general -Startrow 2 -Headername Label,value -Key Label -GridView
        The "General" page has a title and two unlabelled columns with the CPU, Memory, Domain, Disk and so on 
        So this version starts at row 2 to skip the tiltle and labels the first column "label" and the Second "Value"; the label acts as the key
        and the result is display on using grid view. Note that grid view works best when the number of columns is small. 
#>
[cmdletbinding()]
    Param(
        #First file to compare 
        [parameter(Mandatory=$true)]
        $Referencefile ,
        #Second file to compare
        [parameter(Mandatory=$true)]
        $Differencefile   ,
        #Name(s) of worksheets to compare.
        $WorkSheetName   = "Sheet1",
        #Name of a column which is unique and will be used to add a row to the DIFF object, default is "Name" 
        $Key             = "Name" ,
        #Properties to include in the DIFF - supports wildcards, default is "*"
        $Property        = "*"    ,
        #Properties to exclude from the the search - supports wildcards 
        $ExcludeProperty ,
        #Specifies custom property names to use, instead of the values defined in the column headers of the TopRow.
        [Parameter(ParameterSetName='B', Mandatory)]
        [String[]]$Headername,   
        #Automatically generate property names (P1, P2, P3, ..) instead of the using the values the top row of the sheet
        [Parameter(ParameterSetName='C', Mandatory)]
        [switch]$NoHeader, 
        #The row from where we start to import data, all rows above the StartRow are disregarded. By default this is the first row.
        [int]$Startrow, 
        #If specified, highlights the DIFF rows 
        [System.Drawing.Color]$BackgroundColor,
        #If specified identifies the tabs which contain DIFF rows  (ignored if -backgroundColor is omitted)   
        [System.Drawing.Color]$TabColor,
        #If specified, highlights the DIFF columns in rows which have the same key.  
        [System.Drawing.Color]$FontColor,
        #If specified opens the Excel workbooks instead of outputting the diff to the console
        [Switch]$Show, 
        #If specified, tries to the show the DIFF in a gridview. (Works best with few columns) 
        [switch]$GridView
    )

    $oneFile = ((Resolve-Path -Path $Referencefile).path -eq (Resolve-Path -Path $Differencefile).path)

    if ($Key -eq "Name" -and $NoHeader) {$key = "p1"} 

    #If we have one file , we mush have two different worksheet names. If we have two files we can a single string or two strings. 
    if     ($onefile -and ( ($WorkSheetName.count -ne 2) -or $WorkSheetName[0] -eq $WorkSheetName[1] ) ) {
        Write-Warning -Message "If both the Reference and difference file are the same then worksheet name must provide 2 different names" 
        return
    }
    if     ($WorkSheetName.count -eq 2)  {$worksheet1 = $WorkSheetName[0] ; $WorkSheet2 = $WorkSheetName[1]} 
    elseif ($WorkSheetName -is [string]) {$worksheet1 = $WorkSheet2 = $WorkSheetName}
    else   {Write-Warning -Message "You must provide either a single worksheet name or two names." ; return }   
    
    #If the paths are wrong, files are locked or the worksheet names are wrong we won't be able to continue
    $params= @{ ErrorAction = [System.Management.Automation.ActionPreference]::Stop } 
    foreach ($p in @("HeaderName","NoHeader","StartRow")) {if ($PSBoundParameters[$p]) {$params[$p] = $PSBoundParameters[$p]}}
    try   {
        $Sheet1 = Import-Excel -Path $Referencefile  -WorksheetName $WorkSheet1 @params                                                                       
        $Sheet2 = Import-Excel -Path $Differencefile -WorksheetName $WorkSheet2 @Params 
    }
    Catch {Write-Warning -Message "Could not read the worksheet from $Referencefile and/or $Differencefile." ; return } 

    #Get Column headings and create a hash table of Name to column letter. 
    $headings = $Sheet1[-1].psobject.Properties.name # This preserves the sequence - using get-member would sort them alphabetically!
    $Columns  = @{} 
    $i = 65   ; foreach ($h in $headings) {$Columns[$h] = [char]($i ++) }  

    #Make a list of properties headings using the Property (default "*") and ExcludeProperty parameters 
    $PropList = @() 
    foreach ($p in $Property)        {$PropList += ($headings.where({$_ -like    $p}) )} 
    foreach ($p in $ExcludeProperty) {$PropList  =  $PropList.where({$_ -notlike $p})  } 
    $PropList = $PropList | Select-Object -Unique 
    if (($headings -contains $key) -and ($PropList -notcontains $Key)) {$PropList += $Key}
    if ($PropList.Count -eq 0)  {Write-Warning -Message "No Columns are selected with -Property = '$Property' and -excludeProperty = '$ExcludeProperty'." ; return}

    #If we add the row numbes to data and include them in the diff, inserting a row will mean all subsequent rows are different so instead ... 
    #... build hash tables with the "key" column as the key and the row in the spreadsheet where it appears as the value. Row 1 is headers so the first data row is 2 
    $rows1      = @{} ; 
    $rows2      = @{} ;
    if ($PropList -contains $Key) {
        $i = 2  ; foreach ($row in $Sheet1) {$rows1[$row.$key] = ($i ++) } 
        $i = 2  ; foreach ($row in $Sheet2) {$rows2[$row.$key] = ($i ++) } 
    }
    else {Write-Warning -Message "Could not find a column '$key' to use as a key - DIFF rows will not have numbers."}

    #Do the comparison and add file,sheet and row to the result - these are prefixed with "_" to show they are added but the addition still might fail so make sure we have some DIFF
    $diff = Compare-Object $Sheet1 $Sheet2 -Property $PropList  
    $diff = $diff | Select-Object -Property (@(
                @{n="_Side";  e={$_.SideIndicator }} 
                @{n="_File";  e={if ($_.SideIndicator -eq '=>') {$Differencefile} else {$Referencefile } }} ,
                @{n="_Sheet"; e={if ($_.SideIndicator -eq '=>') {$worksheet2 }    else {$worksheet1    } }}   ,
                @{n='_Row';   e={if     ($_.$key -and $_.SideIndicator -eq '=>') {$rows2[$_.$key]} elseif ($_.$key) {$rows1[$_.$key]} else   { "" } }}  
            ) + $PropList)  | Sort-Object -Property row,file

    #if BackgroundColor was specified, set it on extra or extra or changed rows - but remember we we only have row numbers if we have a key 
    if (($PropList -contains $Key) -and $BackgroundColor) {
        #Differences may only exist in one file. So gather the changes for each file; open the file, update each impacted row, save the file  
        $updates =  $diff | Group-object -Property "_File"
        foreach   ($file in $updates) {
            try   {$xl  = Open-ExcelPackage -Path $file.name }
            catch {Write-warning -Message "Can't open $($file.Name) for writing." ; return} 
            foreach ($row in $file.group)  {
                $ws = $xl.Workbook.Worksheets[$row._Sheet]
                $range = $ws.Dimension -replace "\d+",$row._row
                Set-Format -WorkSheet $ws -Range $range -BackgroundColor $BackgroundColor         
            }
            if ($TabColor) {
                foreach ($tab in ($file.group._sheet | Select-Object -Unique)) {
                    $xl.Workbook.Worksheets[$tab].TabColor = $TabColor
                 }
            }
            $xl.save()  ; $xl.Stream.Close() ; $xl.Dispose()
        }
    }

    #if font colour was specified, set it on changed properties where the same key appears in both sheets. 
    if (($PropList -contains $Key) -and $FontColor) {
        $updates = $diff | Group-object -Property $Key | where {$_.count -eq 2} 
        if ($updates) {
            $XL1 = Open-ExcelPackage -path $Referencefile
            if ($oneFile ) {$xl2 = $xl1} 
            else           {$xl2 = Open-ExcelPackage -path $Differencefile }
            foreach ($u in $updates) {
                 foreach ($p in $proplist) {
                    if($u.Group[0].$p -ne $u.Group[1].$p ) {
                        Set-Format -WorkSheet $xl1.Workbook.Worksheets[$u.Group[0]._sheet] -Range ($Columns[$p] + $u.Group[0]._Row) -FontColor $FontColor
                        Set-Format -WorkSheet $xl2.Workbook.Worksheets[$u.Group[1]._sheet] -Range ($Columns[$p] + $u.Group[1]._Row) -FontColor $FontColor
                    } 
                } 
            }
            $xl1.Save()                     ; $xl1.Stream.Close() ; $xl1.Dispose()
            if (-not $oneFile) {$xl2.Save() ; $xl2.Stream.Close() ; $xl2.Dispose()}
        }
    }

    if     ($show)     { 
          Start-Process -FilePath $Referencefile 
          if (-not $oneFile)  { Start-Process -FilePath $Differencefile }
    }
    elseif ($GridView) { 
        if ($StartRow) {$lastrow = $StartRow} else {$lastRow = 1} 
        $diff | Group-Object -Property $key | foreach  {
            $hash = [ordered]@{row = $lastRow; $key = $_.Name; } ; 
            foreach ($row IN $_.Group) {
                if  ($row._Side -eq "=>") {$lastRow = $hash.row =  $row._Row  } 
                foreach ($p in $proplist.Where({$_ -ne $key})) {$hash[($row._Side+$P)] =$row.$P}
        } 
        [Pscustomobject]$hash   }  | Sort-Object -Property row| Update-FirstObjectProperties | Out-GridView -Title "Comparing $Referencefile::$worksheet1 (=>) with $Differencefile::$WorkSheet2 (<=)"   
    }
    else               {return $diff}
}