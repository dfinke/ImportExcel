function Export-MultipleExcelSheets {
    <#
        .Synopsis
        Takes a hash table of scriptblocks and exports each as a sheet in an Excel file    

        .Example
$p = Get-Process

$InfoMap = @{
    PM                 = { $p | Select-Object company, pm }
    Handles            = { $p | Select-Object company, handles }
    Services           = { Get-Service }
}

Export-MultipleExcelSheets -Path $xlfile -InfoMap $InfoMap -Show -AutoSize        
    #>
    param(
        [Parameter(Mandatory = $true)]
        $Path,
        [Parameter(Mandatory = $true)]
        [hashtable]$InfoMap,
        [string]$Password,
        [Switch]$Show,
        [Switch]$AutoSize
    )

    $parameters = @{ } + $PSBoundParameters
    $parameters.Remove("InfoMap")
    $parameters.Remove("Show")

    $parameters.Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

    foreach ($entry in $InfoMap.GetEnumerator()) {
        if ($entry.Value -is [scriptblock]) {
            Write-Progress -Activity "Exporting" -Status "$($entry.Key)"
            $parameters.WorkSheetname = $entry.Key

            & $entry.Value | Export-Excel @parameters
        }
        else {
            Write-Warning "$($entry.Key) not exported, needs to be a scriptblock"
        }
    }

    if ($Show) { Invoke-Item $Path }
}
