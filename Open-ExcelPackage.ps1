Function Open-ExcelPackage  {
<#
.Synopsis
    Returns an Excel Package Object for the specified XLSX file
.Description
    Import-Excel and Export-Excel open an Excel file, carry out their tasks and close it again.
    Sometimes it is necessary to open a file and do other work on it. Open-Excel package allows the file to be opened for these tasks.
    It takes a KillExcel switch to make sure Excel is not holding the file open; a password parameter for existing protected files,
    and a create switch to set-up a new file if no file already exists.
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
        #The password for a protected worksheet, as a [normal] string (not a secure string.)
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
      .Description
        When working with an Excel packaage object the workbook is held in memory and not saved until the Save() method of the package is called.
        Close package saves and disposes of the package object. It can be called with -NoSave to abandon the file without saving, with a new "SaveAs" filename
        with a password to protect the file. And with Show to open it in Excel. -Calculate will try to update the workbook, although not everything can be recalculated
      .Example
        Close-ExcelPackage -show $excel
        $excel holds a package object, this saves the workbook and loads it into Excel.
      .Example
        Close-ExcelPackage -NoSave $excel
        $excel holds a package object, this disposes of it without writing it to disk.
    #>
    [CmdLetBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword","")]
    Param (
    #Package to close.
    [parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [OfficeOpenXml.ExcelPackage]$ExcelPackage,
    #Open the file in Excel.
    [switch]$Show,
    #Abandon the file without saving.
    [Switch]$NoSave,
    #Save file with a new name (ignored if -NoSave Specified).
    $SaveAs,
    #Password to protect the file.
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

