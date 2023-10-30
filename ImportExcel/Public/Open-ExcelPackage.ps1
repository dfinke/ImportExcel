function Open-ExcelPackage {
    [CmdLetBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
    [OutputType([OfficeOpenXml.ExcelPackage])]
    param(
        #The path to the file to open.
        [Parameter(Mandatory = $true)]$Path,
        #If specified, any running instances of Excel will be terminated before opening the file.
        [switch]$KillExcel,
        #The password for a protected worksheet, as a [normal] string (not a secure string).
        [String]$Password,
        #By default Open-ExcelPackage will only opens an existing file; -Create instructs it to create a new file if required.
        [switch]$Create
    )

    if ($KillExcel) {
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
        if ($Password) { $pkgobj = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path , $Password }
        else { $pkgobj = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path }
        if ($pkgobj) {
            foreach ($w in $pkgobj.Workbook.Worksheets) {
                $sb = [scriptblock]::Create(('$this.workbook.Worksheets["{0}"]' -f $w.name))
                try { 
                    Add-Member -InputObject $pkgobj -MemberType ScriptProperty -Name $w.name -Value $sb -ErrorAction Stop
                }
                catch {
                    Write-Warning "Could not add sheet $($w.name) as 'short cut', you need to access it via `$wb.Worksheets['$($w.name)'] "
                }
            }
            return $pkgobj
        }
    }
    else { Write-Warning "Could not find $path" }
}
