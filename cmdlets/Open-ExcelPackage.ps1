Function Open-ExcelPackage  {
<#
.Synopsis
    Returns an ExcelPackage object for the specified XLSX fil.e
.Description
    Import-Excel and Export-Excel open an Excel file, carry out their tasks and close it again.
    Sometimes it is necessary to open a file and do other work on it.
    Open-ExcelPackage allows the file to be opened for these tasks.
    It takes a -KillExcel switch to make sure Excel is not holding the file open;
    a -Password parameter for existing protected files,
    and a -Create switch to set-up a new file if no file already exists.
.Example
    >
    PS> $excel = Open-ExcelPackage -Path "$env:TEMP\test99.xlsx" -Create
    $ws = Add-WorkSheet -ExcelPackage $excel

   This will create a new file in the temp folder if it doesn't already exist.
   It then adds a worksheet - because no name is specified it will use the
   default name of "Sheet1"
.Example
     >
    PS>     $excel  = Open-ExcelPackage -path "$xlPath" -Password $password
    $sheet1 = $excel.Workbook.Worksheets["sheet1"]
    Set-ExcelRange -Range $sheet1.Cells["E1:S1048576"], $sheet1.Cells["V1:V1048576"]  -NFormat ([cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern)
    Close-ExcelPackage $excel -Show

   This will open the password protected file at $xlPath using the password stored
   in $Password. Sheet1 is selected and formatting applied to two blocks of the sheet;
   then the file is and saved and loaded into Excel.
#>
    [CmdLetBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword","")]
    [OutputType([OfficeOpenXml.ExcelPackage])]
    Param (
        #The path to the file to open.
        [Parameter(Mandatory=$true)]$Path,
        #If specified, any running instances of Excel will be terminated before opening the file.
        [switch]$KillExcel,
        #The password for a protected worksheet, as a [normal] string (not a secure string).
        [String]$Password,
        #By default Open-ExcelPackage will only opens an existing file; -Create instructs it to create a new file if required.
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
        if ($Password) {$pkgobj = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path , $Password }
        else           {$pkgobj = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path }
        if ($pkgobj) {
            foreach ($w in $pkgobj.Workbook.Worksheets) {
                $sb = [scriptblock]::Create(('$this.workbook.Worksheets["{0}"]' -f $w.name))
                Add-Member -InputObject $pkgobj -MemberType ScriptProperty -Name $w.name -Value $sb
            }
            return $pkgobj
        }
    }
    else   {Write-Warning "Could not find $path" }
 }

Function Close-ExcelPackage {
    <#
      .Synopsis
        Closes an Excel Package, saving, saving under a new name or abandoning changes and opening the file in Excel as required.
      .Description
        When working with an ExcelPackage object, the Workbook is held in memory and not saved until the .Save() method of the package is called.
        Close-ExcelPackage saves and disposes of the Package object. It can be called with -NoSave to abandon the file without saving, with a new "SaveAs" filename,
        and/or with a password to protect the file. And -Show will open the file in Excel;
        -Calculate will try to update the workbook, although not everything can be recalculated
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

