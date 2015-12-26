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

        $value = $item.$measure
        if($value -is [string] -or $value -is [System.Enum]) {
            $value = 1
        }
        $h.$key+=$value
    }

    foreach ($entry in $h.GetEnumerator()){
        [PSCustomObject]@{
            Name=$entry.key
            $measure=$entry.value
        }
    }
}