function New-ConditionalText {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system State')]
    param(
        #[Parameter(Mandatory=$true)]
        [Alias('ConditionValue')]
        $Text,
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
        $ConditionalType='ContainsText'
    )

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
}