function New-ConditionalText {
    param(
        [Parameter(Mandatory=$true)]
        $Text,
        [System.Drawing.Color]$ConditionalTextColor="DarkRed",
        [System.Drawing.Color]$BackgroundColor="LightPink",
        [OfficeOpenXml.Style.ExcelFillStyle]$PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid,
        [ValidateSet("ContainsText","NotContainsText","BeginsWith","EndsWith")]
        $ConditionalType="ContainsText"        
    )

    $obj = [PSCustomObject]@{
        Text                 = $Text
        ConditionalTextColor = $ConditionalTextColor
        ConditionalType      = $ConditionalType 
        PatternType          = $PatternType 
        BackgroundColor      = $BackgroundColor 
    }

    $obj.pstypenames.Clear()
    $obj.pstypenames.Add("ConditionalText")
    $obj       
}