try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

BarChart (.\TargetData.ps1) "A BarChart"