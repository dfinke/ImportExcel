function Get-Range ($start=0,$stop,$step=1) {
    for ($idx = $start; $idx -lt $stop; $idx+=$step) {$idx}
}