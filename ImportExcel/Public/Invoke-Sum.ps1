function Invoke-Sum {
    [CmdletBinding()]
    param(
        $Data,
        $Dimension,
        $Measure
    )

    if(!$Measure) {$Measure = $Dimension}

    $h=@{}

    foreach ($item in $Data){
        $key=$item.$Dimension

        if(!$key) {$key="[missing]"}

        if(!$h.ContainsKey($key)) {
            $h.$key=[ordered]@{}
        }

        foreach($m in $Measure) {
            $value = $item.$m
            if($value -is [string] -or $value -is [System.Enum]) {
                $value = 1
            }

            $h.$key.$m+=$value
        }
    }

    foreach ($entry in $h.GetEnumerator()){

        $nh=[ordered]@{Name=$entry.key}

        foreach ($item in $entry.value.getenumerator()) {
            $nh.($item.key)=$item.value
        }

        [pscustomobject]$nh
    }
}