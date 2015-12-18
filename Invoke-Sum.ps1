function Invoke-Sum {
    param($data,$dimension,$measure)

    $h=@{}

    foreach ($item in $data){
        $h.$($item.$dimension)+=$item.$measure
    }

    foreach ($entry in $h.GetEnumerator()){
        [PSCustomObject]@{
            Name=$entry.key
            $measure=$entry.value
        }
    }
    
}