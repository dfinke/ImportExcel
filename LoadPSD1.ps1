if((Get-Module -list ImportExcel) -eq $null) {
    Import-Module $PSScriptRoot\ImportExcel.psd1 -force
}