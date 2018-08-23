Function ColorCompletion {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    [System.Drawing.KnownColor].GetFields() | Where-Object {$_.IsStatic -and $_.name -like "$wordToComplete*" } |
        Sort-Object name | ForEach-Object {New-CompletionResult $_.name $_.name
    }
}

if (Get-Command -Name register-argumentCompleter -ErrorAction SilentlyContinue) {
    Register-ArgumentCompleter -CommandName Export-Excel               -ParameterName TitleBackgroundColor   -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName DataBarColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName ForeGroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Add-ConditionalFormatting  -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName AllDataBackgroundColor -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName TabColor               -ScriptBlock $Function:ColorCompletion
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
    Register-ArgumentCompleter -CommandName Set-Format                 -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Format                 -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Format                 -ParameterName BorderColor            -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Format                 -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Column                 -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Column                 -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Column                 -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Row                    -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Row                    -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Row                    -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
}