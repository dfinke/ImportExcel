try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$r = Get-ChildItem C:\WINDOWS\system32 -File

$TotalSettings = @{ 
    Length = "Sum"
    Name = "Count"
    Extension = @{
        # You can create the formula in an Excel workbook first and copy-paste it here
        # This syntax can only be used for the Custom type
        Custom = "=COUNTIF([Extension];`".exe`")"
    }
}

$r | Export-Excel -TableName system32files -TableStyle Medium10 -TotalSettings $TotalSettings -Show