function Set-WorksheetProtection {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingfunctions', '',Justification='Does not change system state')]
    param (
        [Parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet ,
        [switch]$IsProtected,
        [switch]$AllowAll,
        [switch]$BlockSelectLockedCells,
        [switch]$BlockSelectUnlockedCells,
        [switch]$AllowFormatCells,
        [switch]$AllowFormatColumns,
        [switch]$AllowFormatRows,
        [switch]$AllowInsertColumns,
        [switch]$AllowInsertRows,
        [switch]$AllowInsertHyperlinks,
        [switch]$AllowDeleteColumns,
        [switch]$AllowDeleteRows,
        [switch]$AllowSort,
        [switch]$AllowAutoFilter,
        [switch]$AllowPivotTables,
        [switch]$BlockEditObject,
        [switch]$BlockEditScenarios,
        [string]$LockAddress,
        [string]$UnLockAddress
    )

    if     ($PSBoundParameters.ContainsKey('isprotected') -and  $IsProtected -eq $false) {$worksheet.Protection.IsProtected  = $false}
    elseif ($IsProtected) {
        $worksheet.Protection.IsProtected  = $true
        foreach ($ParName in @('AllowFormatCells',
          'AllowFormatColumns', 'AllowFormatRows',
          'AllowInsertColumns', 'AllowInsertRows', 'AllowInsertHyperlinks',
          'AllowDeleteColumns', 'AllowDeleteRows',
          'AllowSort'         , 'AllowAutoFilter', 'AllowPivotTables')) {
               if ($AllowAll -and -not $PSBoundParameters.ContainsKey($Parname)) {$worksheet.Protection.$ParName = $true}
               elseif ($PSBoundParameters[$ParName] -eq $true )                      {$worksheet.Protection.$ParName = $true}
        }
        if ($BlockSelectLockedCells)   {$worksheet.Protection.AllowSelectLockedCells   = $false }
        if ($BlockSelectUnlockedCells) {$worksheet.Protection.AllowSelectUnLockedCells = $false }
        if ($BlockEditObject)          {$worksheet.Protection.AllowEditObject          = $false }
        if ($BlockEditScenarios)       {$worksheet.Protection.AllowEditScenarios       = $false }
    }
    Else {Write-Warning -Message "You haven't said if you want to turn protection off, or on." }

    if ($LockAddress) {
        Set-ExcelRange     -Range $Worksheet.cells[$LockAddress] -Locked
    }
    elseif ($IsProtected) {
        Set-ExcelRange     -Range $Worksheet.Cells -Locked
    }
    if ($UnlockAddress) {
        Set-ExcelRange     -Range $Worksheet.cells[$UnlockAddress] -Locked:$false
    }
}