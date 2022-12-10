function New-ConditionalFormattingIconSet {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Does not change system State')]
    param(
        [Parameter(Mandatory = $true)]
        $Range,
        [ValidateSet("ThreeIconSet", "FourIconSet", "FiveIconSet")]
        $ConditionalFormat,
        [Switch]$Reverse,
        [Switch]$ShowIconOnly
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

        $bp = @{} + $PSBoundParameters

        $obj = [PSCustomObject]@{
            Range        = $Range
            Formatter    = $ConditionalFormat
            IconType     = $bp.IconType
            Reverse      = $Reverse
            ShowIconOnly = $ShowIconOnly
        }

        $obj.pstypenames.Clear()
        $obj.pstypenames.Add("ConditionalFormatIconSet")

        $obj
    }
}