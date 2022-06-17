<#
    .Synopsis
    Chart download stats for modules/scripts published on the PowerShell Gallery
    .Example
    .\Get-ModuleStats.ps1 ImportExcel
#>

param(
    [Parameter(Mandatory=$true)]
    $moduleName,
    [ValidateSet('Column','Bar','Line','Pie')]
    $chartType="Line"
)

$galleryUrl = "https://www.powershellgallery.com/packages/$moduleName"
$nolegend = '-nolegend'
if($chartType -eq 'pie') {$nolegend = $null}
$code = "$($chartType)Chart (Get-HtmlTable $galleryUrl -FirstDataRow 1 | sort lastupdated -desc) -title 'Download stats for $moduleName' $nolegend"

$code | Invoke-Expression