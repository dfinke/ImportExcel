try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$plt = New-Plot
$plt.Plot((Get-Range 0 5 .02|Foreach-Object {[math]::Cos(2*[math]::pi*$_)}))
$plt.SetChartSize(800,300)
$plt.Show()