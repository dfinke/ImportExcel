function Invoke-Sum {
    param($data,$dimension,$measure)

    $h=@{}

    foreach ($item in $data){
        $key=$item.$dimension
        
        if(!$key) {$key="[missing]"}
        
        $h.$key+=$item.$measure
    }

    foreach ($entry in $h.GetEnumerator()){
        [PSCustomObject]@{
            Name=$entry.key
            $measure=$entry.value
        }
    }
    
}