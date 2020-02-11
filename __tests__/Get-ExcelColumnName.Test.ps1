$map = @{
    1024  = 'AMJ'
    2048  = 'BZT'
    3072  = 'DND'
    4096  = 'FAN'
    5120  = 'GNX'
    6144  = 'IBH'
    7168  = 'JOR'
    8192  = 'LCB'
    9216  = 'MPL'
    10240 = 'OCV'
    11264 = 'PQF'
    12288 = 'RDP'
    13312 = 'SQZ'
    14336 = 'UEJ'
    15360 = 'VRT'
    16384 = 'XFD'
}

(Get-ExcelColumnName 26).columnName  | Should be 'Z'
(Get-ExcelColumnName 27).columnName  | Should be 'AA'
(Get-ExcelColumnName 28).columnNamee | Should be 'AB'
(Get-ExcelColumnName 30).columnName  | Should be 'AD'
(Get-ExcelColumnName 48).columnName  | Should be 'AV'

1..16 | ForEach-Object {
    $number = $_ * 1024
    (Get-ExcelColumnName $number).columnName | Should be $map.$number
}
