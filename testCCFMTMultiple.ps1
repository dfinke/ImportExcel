ipmo importexcel -Force

$f = "$pwd\test.xlsx"
rm $f -ErrorAction Ignore

$p=@{
    ConditionalFormat = "ThreeIconSet"
    IconType = "Signs"
    Reverse = $true
}

$RuleHandles = New-ConditionalFormattingIconSet -Address "C:C" @p
$RulePM      = New-ConditionalFormattingIconSet -Address "D:D" @p

ps | 
    Where company | 
    Select Company, Name, Handles, PM | 
    Sort Handles -Descending | 
    Export-Excel $f -Show -AutoSize -ConditionalFormat $RuleHandles,$RulePM