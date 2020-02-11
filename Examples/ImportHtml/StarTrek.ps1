try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

Import-Html "https://en.wikipedia.org/wiki/List_of_Star_Trek:_The_Original_Series_episodes" 2