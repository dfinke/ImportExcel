if(!(Get-Command ig -ErrorAction SilentlyContinue)) {

    "Use ``Install-Module NameIT`` to get the needed module from the gallery to support running this script"

    return
}

$sign=@{sign=@( "+", "-" )}
$location=@{location=@("Atlanta", "Newark", "Washington", "Chicago", "Philadelphia", "Houston", "Phoneix")}

$(1..6 | Foreach-Object {

    $from=$to=""
    while($from -eq $to) {
        $from=ig "[location]" -CustomData $location
        $to=ig "[location]" -CustomData $location
    }

    [double]$a=ig "########"
    [double]$b=ig ".####"
    [double]$c=ig "#######"
    [double]$d=ig "[sign].##" -CustomData $sign
    [double]$e=ig "###"
    [double]$f=ig "[sign]##" -CustomData $sign

    #"{0},{1},{2},{3},{4},{5},{6},{7}" -f $from, $to, $a, $b, $c, $d, $e, $f

    [PSCustomObject][Ordered]@{
        From=$from
        To=$to
        RDollars=$a
        RPercent=$b
        MDollars=$c
        MPercent=$d
        Revenue=$e
        Margin=$f
    }
} | ConvertTo-Csv -NoTypeInformation)  -replace '"','' # | Export-Excel
