$plt = New-Plot
$plt.Plot((Get-Range 0 5 .02|%{[math]::Cos(2*[math]::pi*$_)}))
$plt.SetChartSize(800,300)
$plt.Show()