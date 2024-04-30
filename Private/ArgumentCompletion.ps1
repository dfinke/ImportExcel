function ColorCompletion {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    [System.Drawing.KnownColor].GetFields() | Where-Object {$_.IsStatic -and $_.name -like "$wordToComplete*" } |
        Sort-Object name | ForEach-Object {New-CompletionResult $_.name $_.name
    }
}

function ListFonts {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingEmptyCatchBlock", "")]
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    if (-not $script:FontFamilies) {
        $script:FontFamilies = @("","")
        try {
            $script:FontFamilies = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
        }
        catch {}
    }
    $script:FontFamilies.where({$_ -Gt "" -and $_ -like "$wordToComplete*"} ) | ForEach-Object {
        New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList "'$_'" , $_ ,
        ([System.Management.Automation.CompletionResultType]::ParameterValue) , $_
    }
}

function NumberFormatCompletion {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $numformats = [ordered]@{
        "General"       = "General"      # format ID  0
        "Number"        = "0.00"         # format ID  2
        "Percentage"    = "0.00%"        # format ID 10
        "Scientific"    = "0.00E+00"     # format ID 11
        "Fraction"      = "# ?/?"        # format ID 12
        "Short Date"    = "Localized"    # format ID 14 - will be translated to "mm-dd-yy"     which is localized on load by Excel.
        "Short Time"    = "Localized"    # format ID 20 - will be translated to "h:mm"         which is localized on load by Excel.
        "Long Time"     = "Localized"    # format ID 21 - will be translated to "h:mm:ss"      which is localized on load by Excel.
        "Date-Time"     = "Localized"    # format ID 22 - will be translated to "m/d/yy h:mm"  which is localized on load by Excel.
        "Currency"      = [cultureinfo]::CurrentCulture.NumberFormat.CurrencySymbol + "#,##0.00"
        "Text"          = "@"              # format ID 49
        "h:mm AM/PM"    = "h:mm AM/PM"     # format ID 18
        "h:mm:ss AM/PM" = "h:mm:ss AM/PM"  # format ID 19
        "mm:ss"         = "mm:ss"          # format ID 45
        "[h]:mm:ss"     = "Elapsed hours"  # format ID 46
        "mm:ss.0"       = "mm:ss.0"        # format ID 47
        "d-mmm-yy"      = "Localized"      # format ID 15 which is localized on load by Excel.
        "d-mmm"         = "Localized"      # format ID 16 which is localized on load by Excel.
        "mmm-yy"        = "mmm-yy"         # format ID 17 which is localized on load by Excel.
        "0"             = "Whole number"                       # format ID  1
        "0.00"          = "Number, 2 decimals"                 # format ID  2 or "number"
        "#,##0"         = "Thousand separators"                # format ID  3
        "#,##0.00"      = "Thousand separators and 2 decimals" # format ID  4
        "#,"            = "Whole thousands"
        "#.0,,"         = "Millions, 1 Decimal"
        "0%"            = "Nearest whole percentage"           # format ID  9
        "0.00%"         = "Percentage with decimals"           # format ID 10 or "Percentage"
        "00E+00"        = "Scientific"                         # format ID 11 or "Scientific"
        "# ?/?"         = "One Digit fraction"                 # format ID 12 or "Fraction"
        "# ??/??"       = "Two Digit fraction"                 # format ID 13
        "@"             = "Text"                               # format ID 49 or "Text"
    }
    $numformats.keys.where({$_ -like "$wordToComplete*"} ) | ForEach-Object {
        New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList "'$_'" , $_ ,
        ([System.Management.Automation.CompletionResultType]::ParameterValue) , $numformats[$_]
    }
}

function WorksheetArgumentCompleter {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $xlPath = $fakeBoundParameter['Path']
    if (Test-Path -Path $xlPath) {
        $xlSheet = Get-ExcelSheetInfo -Path $xlPath
        $WorksheetNames = $xlSheet.Name
        $WorksheetNames.where( { $_ -like "*$wordToComplete*" }) | foreach-object {
            New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList "'$_'",
            $_ , ([System.Management.Automation.CompletionResultType]::ParameterValue) , $_
        }
    }
}

if   (Get-Command -ErrorAction SilentlyContinue -name Register-ArgumentCompleter) {
    Register-ArgumentCompleter -CommandName Export-Excel               -ParameterName TitleBackgroundColor   -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName DataBarColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName ForeGroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName AllDataBackgroundColor -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName TabColor               -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Join-Worksheet             -ParameterName TitleBackgroundColor   -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Merge-Worksheet            -ParameterName AddBackgroundColor     -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Merge-Worksheet            -ParameterName ChangeBackgroundColor  -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Merge-Worksheet       `    -ParameterName DeleteBackgroundColor  -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Merge-MulipleSheets        -ParameterName KeyFontColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Merge-MulipleSheets        -ParameterName AddBackgroundColor     -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Merge-MulipleSheets        -ParameterName ChangeBackgroundColor  -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Merge-MulipleSheets   `    -ParameterName DeleteBackgroundColor  -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Merge-MulipleSheets        -ParameterName KeyFontColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName New-ConditionalText        -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName New-ConditionalText        -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName New-ConditionalText        -ParameterName ConditionalTextColor   -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName New-ExcelStyle             -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName New-ExcelStyle             -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName New-ExcelStyle             -ParameterName BorderColor            -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName New-ExcelStyle             -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRange             -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRange             -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRange             -ParameterName BorderColor            -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRange             -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelColumn            -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelColumn            -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelColumn            -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRow               -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRow               -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRow               -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName New-ExcelStyle             -ParameterName FontName               -ScriptBlock $Function:ListFonts
    Register-ArgumentCompleter -CommandName Set-ExcelColumn            -ParameterName FontName               -ScriptBlock $Function:ListFonts
    Register-ArgumentCompleter -CommandName Set-ExcelRange             -ParameterName FontName               -ScriptBlock $Function:ListFonts
    Register-ArgumentCompleter -CommandName Set-ExcelRow               -ParameterName FontName               -ScriptBlock $Function:ListFonts
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName NumberFormat           -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Export-Excel               -ParameterName NumberFormat           -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName New-ExcelStyle             -ParameterName NumberFormat           -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRange             -ParameterName NumberFormat           -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelColumn            -ParameterName NumberFormat           -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Set-ExcelRow               -ParameterName NumberFormat           -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Add-PivotTable             -ParameterName PivotNumberFormat      -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName New-PivotTableDefinition   -ParameterName PivotNumberFormat      -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName New-ExcelChartDefinition   -ParameterName XAxisNumberformat      -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName New-ExcelChartDefinition   -ParameterName YAxisNumberformat      -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Add-ExcelChart             -ParameterName XAxisNumberformat      -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Add-ExcelChart             -ParameterName YAxisNumberformat      -ScriptBlock $Function:NumberFormatCompletion
    Register-ArgumentCompleter -CommandName Import-Excel               -ParameterName WorksheetName          -ScriptBlock $Function:WorksheetArgumentCompleter
}
