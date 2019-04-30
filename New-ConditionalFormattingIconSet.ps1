function New-ConditionalFormattingIconSet {
    <#
      .SYNOPSIS
        Creates an object which describes a conditional formatting rule a for 3,4 or 5 icon set.
      .DESCRIPTION
        Export-Excel takes a -ConditionalFormat parameter which can hold one or more descriptions for conditional formats;
        this command builds the defintion of a Conditional formatting rule for an icon set.
      .PARAMETER Range
        The range of cells that the conditional format applies to.
      .PARAMETER ConditionalFormat
        The type of rule: one of "ThreeIconSet","FourIconSet" or "FiveIconSet"
      .PARAMETER IconType
        The name of an iconSet - different icons are available depending on whether 3,4 or 5 icon set is selected.
      .PARAMETER Reverse
        Use the icons in the reverse order.
      .Example
        $cfRange = [OfficeOpenXml.ExcelAddress]::new($topRow, $column, $lastDataRow, $column)
        $cfdef = New-ConditionalFormattingIconSet -Range $cfrange -ConditionalFormat ThreeIconSet -IconType Arrows
        Export-Excel -ExcelPackage $excel -ConditionalFormat $cfdef -show

        The first line creates a range - one column wide in the column $column, running
        from $topRow to $lastDataRow.
        The second line creates a definition object using this range
        and the third uses Export-Excel with an open package to apply the format and
        save and open the file.
      .Link
        Add-Add-ConditionalFormatting
        New-ConditionalText
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system State')]
    param(
        [Parameter(Mandatory=$true)]
        $Range,
        [ValidateSet("ThreeIconSet","FourIconSet","FiveIconSet")]
        $ConditionalFormat,
        [Switch]$Reverse
    )

    DynamicParam {
        $IconType = New-Object System.Management.Automation.ParameterAttribute
        $IconType.Position = 2
        $IconType.Mandatory = $true

        $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]

        $attributeCollection.Add($IconType)

        switch ($ConditionalFormat) {
            "ThreeIconSet" {
                $IconTypeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('IconType', [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting3IconsSetType], $attributeCollection)
            }

            "FourIconSet" {
                $IconTypeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('IconType', [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting4IconsSetType], $attributeCollection)
            }

            "FiveIconSet" {
                $IconTypeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('IconType', [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting5IconsSetType], $attributeCollection)
            }
        }

        $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

        $paramDictionary.Add('IconType', $IconTypeParam)

        return $paramDictionary
    }

    End {

        $bp = @{}+$PSBoundParameters

        $obj = [PSCustomObject]@{
            Range     = $Range
            Formatter = $ConditionalFormat
            IconType  = $bp.IconType
            Reverse   = $Reverse
        }

        $obj.pstypenames.Clear()
        $obj.pstypenames.Add("ConditionalFormatIconSet")

        $obj
    }
}