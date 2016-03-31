function Invoke-Sum {
    param(
        $data,
        $dimension,
        $measure
    )

    if(!$measure) {$measure = $dimension}

    $h=@{}

    foreach ($item in $data){
        $key=$item.$dimension

        if(!$key) {$key="[missing]"}

        if(!$h.ContainsKey($key)) {
            $h.$key=[ordered]@{}
        }

        foreach($m in $measure) {        
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