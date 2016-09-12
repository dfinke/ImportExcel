function New-ConditionalText {
    param(
        #[Parameter(Mandatory=$true)]
        $Text,
        [System.Drawing.Color]$ConditionalTextColor="DarkRed",
        [System.Drawing.Color]$BackgroundColor="LightPink",
        [String]$Range,
        [OfficeOpenXml.Style.ExcelFillStyle]$PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid,
        [ValidateSet(
       	    "LessThan","LessThanOrEqual","GreaterThan","GreaterThanOrEqual",
            "NotEqual","Equal","ContainsText","NotContainsText","BeginsWith","EndsWith",
            "Last7Days","LastMonth","LastWeek",
            "NextMonth","NextWeek",
            "ThisMonth","ThisWeek",
            "Today","Tomorrow","Yesterday",
            "DuplicateValues",
            "AboveOrEqualAverage","BelowAverage","AboveAverage",
            "Top", "TopPercent", "ContainsBlanks"
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