try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

ColumnChart (.\TargetData.ps1) "A ColumnChart"
