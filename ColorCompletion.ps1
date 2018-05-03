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
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName AllDataBackgroundColor -ScriptBlock $Function:ColorCompletion    
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Compare-Worksheet          -ParameterName TabColor               -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Format                 -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Format                 -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Format                 -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Column                 -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Column                 -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Column                 -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Row                    -ParameterName BackgroundColor        -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Row                    -ParameterName FontColor              -ScriptBlock $Function:ColorCompletion
    Register-ArgumentCompleter -CommandName Set-Row                    -ParameterName PatternColor           -ScriptBlock $Function:ColorCompletion
}