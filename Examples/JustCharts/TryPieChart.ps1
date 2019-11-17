try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

PieChart (.\TargetData.ps1) "A PieChart"