function New-PSItem {

    $totalArgs = $args.Count

    if($args[-1] -is [array]) {
        $script:PSItemHeader=$args[-1]
        $totalArgs-=1
    }

    $h=[ordered]@{}

    for ($idx = 0; $idx -lt $totalArgs; $idx+=1) {
        if($PSItemHeader) {
            $key = $PSItemHeader[$idx]
        } else {
            $key = "P$($idx+1)"
        }

        $h.$key=$args[$idx]
    }

    [PSCustomObject]$h
}