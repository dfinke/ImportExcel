<#
    .Synopsis
    Chart download stats for modules/scripts published on the PowerShell Gallery
    .Example
    .\Get-ModuleStats.ps1 ImportExcel
#>

param(
    $moduleName = "ImportExcel",
    [ValidateSet('Column','Bar','Line','Pie')]
    $chartType="Line"
)

$download = Get-HtmlTable "https://www.powershellgallery.com/packages/$moduleName" -FirstDataRow 1 |
        Select-Object  @{n="Version";e={$v = $Null ; if    ($_.version -is [valuetype]) {[string][version]($_.version.tostring("0.0")) }
                                                    elseif ($_.version -is [string] -and [version]::TryParse($_.version.trim(),[ref]$v))  {$v}
                                                    else   {$_.Version.trim() -replace "\s+"," " } }},
                        Downloads, @{n="LastUpdated";e={[datetime]$_.last_updated}} |
            Sort-Object lastupdated -Descending

& "$($chartType)Chart"  $download "Download stats for $moduleName" -nolegend:($chartype -ne 'pie')
