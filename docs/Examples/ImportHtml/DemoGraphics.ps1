try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}


Import-Html "http://en.wikipedia.org/wiki/Demographics_of_India" 4