Function ConvertTo-ExcelXlsx {
[CmdletBinding()]
PARAM
(
    [parameter(Mandatory=$true, ValueFromPipeline)]
    [ValidateScript({
        if(-Not ($_ | Test-Path) ){
            throw "File not found" 
        }
        if(-Not ($_ | Test-Path -PathType Leaf) ){
            throw "Folder paths are not allowed"
        }
        return $true
    })]
    [string]$Path,
    [parameter(Mandatory=$false)]
    [switch]$Force
)
    $xlFixedFormat = 51 #Constant for XLSX Workbook
    $xlsFile = Get-Item -Path $Path
    $xlsxPath = "{0}x" -f $xlsFile.FullName

    if($xlsFile.Extension -ne ".xls"){
        throw "Expected .xls extension"
    }

    if(Test-Path -Path $xlsxPath){
        if($Force){
            Remove-Item $xlsxPath -Force
            Write-Verbose $("Removed {0}" -f $xlsxPath)
        } else {
            throw "{0} already exists!" -f $xlsxPath
        }
    }
    
    try{    
        $Excel = New-Object -ComObject "Excel.Application"
    } catch {
        throw "Could not create Excel.Application ComObject. Please verify that Excel is installed."
    }

    $Excel.Visible = $false
    $Excel.Workbooks.Open($xlsFile.FullName) | Out-Null
    $Excel.ActiveWorkbook.SaveAs($xlsxPath, $xlFixedFormat)
    $Excel.ActiveWorkbook.Close()
    $Excel.Quit()
}

