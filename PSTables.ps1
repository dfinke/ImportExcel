function Format-TransposeTable {
    [CmdletBinding()]
    param (
        [Parameter(
            Position = 0,
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )] [object[]]$Object
    )
    begin { $i = 0; }

    process {
        foreach ($myObject in $Object) {
            if ($myObject.GetType().Name -eq 'hashtable' -or $myObject.GetType().Name -eq 'OrderedDictionary') {
                Write-Verbose " Format-TransposeTable - Converting HashTable/OrderedDictionary to PSCustomObject - $($myObject.GetType().Name)"
                $output = New-Object -TypeName PsObject;
                Add-Member -InputObject $output -MemberType ScriptMethod -Name AddNote -Value {
                    Add-Member -InputObject $this -MemberType NoteProperty -Name $args[0] -Value $args[1];
                };
                $myObject.Keys | Sort-Object | % {
                    $output.AddNote($_, $myObject.$_);
                }
                $output;
            } else {
                Write-Verbose " Format-TransposeTable - Converting PSCustomObject to HashTable/OrderedDictionary - $($myObject.GetType().Name)"
                # Write-Warning "Index $i is not of type [hashtable]";
                $output = [ordered] @{};
                $myObject | Get-Member -MemberType *Property | % {
                    $output.($_.name) = $myObject.($_.name);
                }
                $output

            }
            $i += 1;
        }
    }
}

function Format-PSTableConvertType3 {
    [CmdletBinding()]
    param (
        $Object,
        [switch] $SkipTitles

    )
    Write-Verbose 'Format-PSTableConvertType3 - Option 3'
    $Array = New-ArrayList
    ### Add Titles
    if (-not $SkipTitles) {
        $Titles = New-ArrayList
        Add-ToArray -List $Titles -Element 'Name'
        Add-ToArray -List $Titles -Element 'Value'
        Add-ToArray -List $Array -Element $Titles
    }
    ### Add Data
    foreach ($O in $Object) {
        foreach ($Key in $O.Keys) {
            # Write-Verbose "Test2 - $Key - $($O[$Key])"
            $ArrayValues = New-ArrayList
            Add-ToArray -List $ArrayValues -Element $Key
            Add-ToArray -List $ArrayValues -Element $O[$Key]
            Add-ToArray -List $Array -Element $ArrayValues
        }
    }
    return , $Array
}
function Format-PSTableConvertType2 {
    [CmdletBinding()]
    param(
        $Object,
        [switch] $SkipTitles
    )
    #Write-Verbose 'Format-PSTableConvertType2 - Option 2'
    $Array = New-ArrayList
    ### Add Titles
    if (-not $SkipTitle) {
        # Write-Verbose 'Format-PSTableConvertType2 - Option 2 - Titles Start'
        $Titles = New-ArrayList
        foreach ($O in $Object) {
            foreach ($Name in $O.PSObject.Properties.Name) {
                #Write-Verbose "my title is $Name"
                Add-ToArray -List $Titles -Element $Name
            }
            break
        }
        Add-ToArray -List ($Array) -Element $Titles
        # Write-Verbose 'Format-PSTableConvertType2 - Option 2 - Titles End'
    }
    ### Add Data
    #Write-Verbose 'Format-PSTableConvertType2 - Option 2 - Data Start'
    foreach ($O in $Object) {
        $ArrayValues = New-ArrayList
        foreach ($Name in $O.PSObject.Properties.Name) {
            #Write-Verbose "my name is $Value"

            Add-ToArray -List $ArrayValues -Element  $O.$Name
        }
        Add-ToArray -List $Array -Element $ArrayValues
    }
    #Write-Verbose 'Format-PSTableConvertType2 - Option 2 - Data End'
    return , $Array
}
function Format-PSTableConvertType1 {
    [CmdletBinding()]
    param (
        $Object,
        [switch] $SkipTitles
    )
    Write-Verbose 'Format-PSTableConvertType1 - Option 1'
    $Array = New-ArrayList
    ### Add Titles
    if (-not $SkipTitles) {
        $Titles = New-ArrayList
        Add-ToArray -List $Titles -Element 'Name'
        Add-ToArray -List $Titles -Element 'Value'
        Add-ToArray -List $Array -Element $Titles
    }
    ### Add Data
    foreach ($Key in $Object.Keys) {
        Write-Verbose "$Key"
        Write-Verbose "$Object.$Key"
        #$Array += , @($Key, $Object.$Key)
        $ArrayValues = New-ArrayList
        Add-ToArray -List $ArrayValues -Element $Key
        Add-ToArray -List $ArrayValues -Element $Object.$Key
        Add-ToArray -List $Array -Element $ArrayValues
    }

    return , $Array
}
function Format-PSTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] $Object,
        [switch] $SkipTitle
    )

    $Type = Get-ObjectType -Object $Object
    Write-Verbose "Format-PSTable - Type: $($Type.ObjectTypeName)"
    if ($Type.ObjectTypeName -eq 'Object[]' -or
        $Type.ObjectTypeName -eq 'Object' -or $Type.ObjectTypeName -eq 'PSCustomObject' -or
        $Type.ObjectTypeName -eq 'Collection`1') {
        #Write-Verbose 'Level 0-0'
        if ($Type.ObjectTypeInsiderName -eq 'string') {
            #Write-Verbose 'Level 1-0'
            return Format-PSTableConvertType1 -Object $Object -SkipTitle:$SkipTitle
        } elseif ($Type.ObjectTypeInsiderName -eq 'Object' -or $Type.ObjectTypeInsiderName -eq 'PSCustomObject') {
            # Write-Verbose 'Level 1-1'
            return Format-PSTableConvertType2 -Object $Object -SkipTitle:$SkipTitle
        } elseif ($Type.ObjectTypeInsiderName -eq 'HashTable' -or $Type.ObjectTypeInsiderName -eq 'OrderedDictionary' ) {
            # Write-Verbose 'Level 1-2'
            return Format-PSTableConvertType3 -Object $Object -SkipTitle:$SkipTitle
        } else {
            # Covers ADDriveInfo and other types of objects
            # Write-Verbose 'Level 1-3'
            return Format-PSTableConvertType2 -Object $Object -SkipTitle:$SkipTitle
        }
    } elseif ($Type.ObjectTypeName -eq 'HashTable' -or $Type.ObjectTypeName -eq 'OrderedDictionary' ) {
        #Write-Verbose 'Level 0-1'
        return Format-PSTableConvertType3 -Object $Object -SkipTitle:$SkipTitle
    } else {
        #Write-Verbose 'Level 0-2'
        # Covers ADDriveInfo and other types of objects
        return Format-PSTableConvertType2 -Object $Object -SkipTitle:$SkipTitle
    }
    throw 'Not supported? Weird'
}
function Show-TableVisualization {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] $Object,
        [switch] $Color
    )
    if ($Color) { Write-Color "[i] This is how table looks like in Format-Table" -Color Yellow }
    Write-Verbose '[i] This is how table looks like in Format-Table'
    $Object | Format-Table -AutoSize
    $Data = Format-PSTable $Object #-Verbose
    # $Data.Count
    #$Data[0].Count
    if ($Color) { Write-Color "[i] Rows Count ", $Data.Count, " Column Count ", $Data[0].Count -Color Yellow }
    Write-Verbose "[i] Rows Count $($Data.Count) Column Count $($Data[0].Count)"
    $RowNr = 0
    if ($Color) { Write-Color "[i] Presenting table after conversion" -Color Yellow }
    foreach ($Row in $Data) {
        $ColumnNr = 0
        foreach ($Column in $Row) {
            if ($Color) { Write-Color 'Row: ', $RowNr, ' Column: ', $ColumnNr, " Data: ", $Column -Color White, Yellow, White, Green }
            Write-Verbose "Row: $RowNr Column: $ColumnNr Data: $Column"
            $ColumnNr++
        }
        $RowNr++
    }
}