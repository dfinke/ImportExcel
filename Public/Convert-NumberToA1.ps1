Function Convert-NumberToA1 {
  Param([parameter(Mandatory=$true)]
        [decimal]$number)
  $number = $number -replace "\..*",""
  $a1Value = $null
  While ($number -gt 0) {
    [decimal]$multiplier = [system.math]::Floor(($number / 26))
    [int]$charNumber = $number - ($multiplier * 26)
    If ($charNumber -eq 0) { $multiplier-- ; $charNumber = 26 }
    $a1Value = [char]($charNumber + 64) + $a1Value
    $number = $multiplier
  }
  Return $a1Value
}
