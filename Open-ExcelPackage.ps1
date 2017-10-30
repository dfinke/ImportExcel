Function Open-ExcelPackage  {
<#
.Synopsis 
    Returns an Excel Package Object with for the specified XLSX ile 
.Example
    $excel  = Open-ExcelPackage -path $xlPath 
    $sheet1 = $excel.Workbook.Worksheets["sheet1"] 
    set-Format -Address $sheet1.Cells["E1:S1048576"], $sheet1.Cells["V1:V1048576"]  -NFormat ([cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern) 
    close-ExcelPackage $excel -Show

   This will open the file at $xlPath, select sheet1 apply formatting to two blocks of the sheet and close the package 
#>
    [OutputType([OfficeOpenXml.ExcelPackage])]
    Param ([Parameter(Mandatory=$true)]$path,
           [switch]$KillExcel)

        if($KillExcel)         {
            Get-Process -Name "excel" -ErrorAction Ignore | Stop-Process
            while (Get-Process -Name "excel" -ErrorAction Ignore) {}
        }

        $Path          = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        if (Test-Path $path) {New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path }
        Else                 {Write-Warning "Could not find $path" } 
 }

Function Close-ExcelPackage {
<#
.Synopsis 
    Closes an Excel Package, saving, saving under a new name or abandoning changes and opening the file as required 
#>
    Param (
    #File to close
    [parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [OfficeOpenXml.ExcelPackage]$ExcelPackage,
    #Open the file
    [switch]$Show, 
    #Abandon the file without saving
    [Switch]$NoSave,
    #Save file with a new name (ignored if -NoSaveSpecified)
    $SaveAs
    )
    if ( $NoSave)      {$ExcelPackage.Dispose()}
    else {
          if ($SaveAs) {$ExcelPackage.SaveAs( $SaveAs ) } 
          Else         {$ExcelPackage.Save(); $SaveAs = $ExcelPackage.File.FullName }
          $ExcelPackage.Dispose() 
          if ($show)   {Start-Process -FilePath $SaveAs } 
    }
}