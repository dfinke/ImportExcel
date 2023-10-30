try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

Import-Html "http://www.science.co.il/PTelements.asp" 1