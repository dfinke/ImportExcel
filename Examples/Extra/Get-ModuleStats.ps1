param(
    [Parameter(Mandatory=$true)]
    $moduleName,
    [ValidateSet('Column','Bar','Line','Pie')]
    $chartType="Line"    
)

$galleryUrl = "https://www.powershellgallery.com/packages/$moduleName"
$nolegend = '-nolegend'
if($chartType -eq 'pie') {$nolegend = $null}
$code = "$($chartType)Chart (Get-HtmlTable $galleryUrl 0) -title 'Download stats for $moduleName' $nolegend"

$code | Invoke-Expression