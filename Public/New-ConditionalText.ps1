function New-ConditionalText {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system State')]
    param(
        #[Parameter(Mandatory=$true)]
        [Alias('ConditionValue')]
        $Text,
        $Comment,
        [Alias('ForeGroundColor')]
        $ConditionalTextColor=[System.Drawing.Color]::DarkRed,
        $BackgroundColor=[System.Drawing.Color]::LightPink,
        [String]$Range,
        [OfficeOpenXml.Style.ExcelFillStyle]$PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid,
        [ValidateSet(
            'LessThan',        'LessThanOrEqual',      'GreaterThan',    'GreaterThanOrEqual',
            'Equal',           'NotEqual',
            'Top',             'TopPercent',           'Bottom',         'BottomPercent',
            'ContainsText',    'NotContainsText',      'BeginsWith',     'EndsWith',
            'ContainsBlanks',  'NotContainsBlanks',    'ContainsErrors', 'NotContainsErrors',
            'DuplicateValues', 'UniqueValues',
            'Tomorrow',        'Today',                'Yesterday',      'Last7Days',
            'NextWeek',        'ThisWeek',             'LastWeek',
            'NextMonth',       'ThisMonth',            'LastMonth',
            'AboveAverage',    'AboveOrEqualAverage',  'BelowAverage',  'BelowOrEqualAverage',
            'Expression'
        )]
        [Alias('RuleType')]
        $ConditionalType='ContainsText',
        $Array,
        $Column,
        [Switch]$NoWarning
    )
    $CommandName = $MyInvocation.MyCommand

    If (-not$Array -and -not$NoWarning) {Write-Warning "$CommandName, Array is empty, column ""$Column"" and Text ""$Text"", Comment: ""$Comment"""}
    If ($Column -eq "" -and -not$NoWarning) {Write-Warning "$CommandName, Column is empty, Text ""$Text"", Comment: ""$Comment"""}

    #region determine Range based on Column
    If ($Array -and $Column -ne "") {
        $Columns = $Array[0].psobject.Properties | foreach { $_.Name }
        $Range = ""
        Foreach ($Col in $Columns) {
            #Write-Verbose "Worksheet:$Worksheet`tColumn:$Column`tRange: $Range"
            If ("$Col" -eq "$Column") {
                $Iteration = [array]::IndexOf($Columns, $Column)
                $ColName = Convert-NumberToA1 ($Iteration + 1)
                $Range = "$ColName"+":"+"$ColName"
                #Write-Verbose "Worksheet:$Worksheet`tColumn:$Column`tRange: $Range"
            } # This is the Column we are looking for
        }

        If (-not$Range -and -not$NoWarning) {Write-Warning "$CommandName, Column ""$Column"" not found, Text ""$Text"", Comment: ""$Comment"""}
        If ($Range) {

            $obj = [PSCustomObject]@{
                Text                 = $Text
                ConditionalTextColor = $ConditionalTextColor
                ConditionalType      = $ConditionalType
                PatternType          = $PatternType
                Range                = $Range
                BackgroundColor      = $BackgroundColor
            }
            $obj.pstypenames.Clear()
            $obj.pstypenames.Add("ConditionalText")
            $obj
        } # If ($Range)
    } # If ($Array -and $Column -ne "")
    #endregion determine Range based on Column  
}
