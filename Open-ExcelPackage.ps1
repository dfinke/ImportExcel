Function Open-ExcelPackage  {
<#
.Synopsis
    Returns an Excel Package Object with for the specified XLSX ile
.Example
    $excel  = Open-ExcelPackage -path $xlPath
    $sheet1 = $excel.Workbook.Worksheets["sheet1"]
    Set-Format -Address $sheet1.Cells["E1:S1048576"], $sheet1.Cells["V1:V1048576"]  -NFormat ([cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern)
    Close-ExcelPackage $excel -Show

   This will open the file at $xlPath, select sheet1 apply formatting to two blocks of the sheet and save the package, and launch it in Excel.
#>
    [CmdLetBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword","")]
    [OutputType([OfficeOpenXml.ExcelPackage])]
    Param (
        #The Path to the file to open
        [Parameter(Mandatory=$true)]$Path,
        #If specified, any running instances of Excel will be terminated before opening the file.
        [switch]$KillExcel,
        [String]$Password,
        #By  default open only opens an existing file; -Create instructs it to create a new file if required.
        [switch]$Create
    )

    if($KillExcel)         {
        Get-Process -Name "excel" -ErrorAction Ignore | Stop-Process
        while (Get-Process -Name "excel" -ErrorAction Ignore) {}
    }

    $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    #If -Create was not specified only open the file if it exists already (send a warning if it doesn't exist).
    if ($Create -and -not (Test-Path -Path $path)) {
        #Create the directory if required.
        $targetPath = Split-Path -Parent -Path $Path
        if (!(Test-Path -Path $targetPath)) {
                Write-Debug "Base path $($targetPath) does not exist, creating"
                $null = New-item -ItemType Directory -Path $targetPath -ErrorAction Ignore
        }
        New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path
    }
    elseif (Test-Path -Path $path) {
        if ($Password) {New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path , $Password }
        else           {New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path }
    }
    else   {Write-Warning "Could not find $path" }
 }

Function Close-ExcelPackage {
<#
.Synopsis
    Closes an Excel Package, saving, saving under a new name or abandoning changes and opening the file in Excel as required.
#>
    [CmdLetBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword","")]
    Param (
    #File to close.
    [parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [OfficeOpenXml.ExcelPackage]$ExcelPackage,
    #Open the file.
    [switch]$Show,
    #Abandon the file without saving.
    [Switch]$NoSave,
    #Save file with a new name (ignored if -NoSave Specified).
    $SaveAs,
    [ValidateNotNullOrEmpty()]
    [String]$Password,
    #Attempt to recalculation the workbook before saving
    [switch]$Calculate
    )
    if ( $NoSave)      {$ExcelPackage.Dispose()}
    else {
          if ($Calculate) {
            try   { [OfficeOpenXml.CalculationExtension]::Calculate($ExcelPackage.Workbook) }
            Catch { Write-Warning "One or more errors occured while calculating, save will continue, but there may be errors in the workbook."}
          }
          if ($SaveAs) {
              $SaveAs = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SaveAs)
              if ($Password) {$ExcelPackage.SaveAs( $SaveAs, $Password ) }
              else           {$ExcelPackage.SaveAs( $SaveAs)}
          }
          Else         {
              if ($Password) {$ExcelPackage.Save($Password) }
              else           {$ExcelPackage.Save()          }
              $SaveAs = $ExcelPackage.File.FullName
          }
          $ExcelPackage.Dispose()
          if ($Show)   {Start-Process -FilePath $SaveAs }
    }
}

