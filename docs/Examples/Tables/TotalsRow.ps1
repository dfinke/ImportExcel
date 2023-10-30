try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$r = Get-ChildItem C:\WINDOWS\system32 -File

$TotalSettings = @{ 
    Name = "Count"
    # You can create the formula in an Excel workbook first and copy-paste it here
    # This syntax can only be used for the Custom type
    Extension = "=COUNTIF([Extension];`".exe`")"
    Length = @{
        Function = "=SUMIF([Extension];`".exe`";[Length])"
        Comment = "Sum of all exe sizes"
    }
}

$r | Export-Excel -TableName system32files -TableStyle Medium10 -TableTotalSettings $TotalSettings -Show