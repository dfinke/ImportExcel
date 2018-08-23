function New-ConditionalText {
    param(
        #[Parameter(Mandatory=$true)]
        $Text,
        [System.Drawing.Color]$ConditionalTextColor="DarkRed",
        [System.Drawing.Color]$BackgroundColor="LightPink",
        [String]$Range,
        [OfficeOpenXml.Style.ExcelFillStyle]$PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid,
        [ValidateSet(
            "LessThan",        "LessThanOrEqual",      "GreaterThan",    "GreaterThanOrEqual",
            "Equal",           "NotEqual",
            "Top",             "TopPercent",           "Bottom",         "BottomPercent",
            "ContainsText",    "NotContainsText",      "BeginsWith",     "EndsWith",
            "ContainsBlanks",  "NotContainsBlanks",    "ContainsErrors", "NotContainsErrors",
            "DuplicateValues", "UniqueValues",
            "Tomorrow",        "Today",                "Yesterday",      "Last7Days",
            "NextWeek",        "ThisWeek",             "LastWeek",
            "NextMonth",       "ThisMonth",            "LastMonth",
            "AboveAverage",    "AboveOrEqualAverage",  "BelowAverage",  "BelowOrEqualAverage"
        )]
        $ConditionalType="ContainsText"
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