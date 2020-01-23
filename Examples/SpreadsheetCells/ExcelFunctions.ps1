try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

Remove-Item "$env:temp\functions.xlsx" -ErrorAction SilentlyContinue

$(
    New-PSItem =2%/12 60 500000 "=pmt(rate,nper,pv)" @("rate", "nper", "pv", "pmt")
    New-PSItem =3%/12 60 500000 "=pmt(rate,nper,pv)"
    New-PSItem =4%/12 60 500000 "=pmt(rate,nper,pv)"
    New-PSItem =5%/12 60 500000 "=pmt(rate,nper,pv)"
    New-PSItem =6%/12 60 500000 "=pmt(rate,nper,pv)"
    New-PSItem =7%/12 60 500000 "=pmt(rate,nper,pv)"
) | Export-Excel "$env:temp\functions.xlsx" -AutoNameRange -AutoSize -Show