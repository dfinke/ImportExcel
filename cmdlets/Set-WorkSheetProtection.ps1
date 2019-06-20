Function Set-WorkSheetProtection {
    [Cmdletbinding()]
    <#
      .Synopsis
        Sets protection on the worksheet
      .Description

      .Example
        Set-WorkSheetProtection -WorkSheet $planSheet -IsProtected -AllowAll -AllowInsertColumns:$false -AllowDeleteColumns:$false -UnLockAddress "A:N"
        Turns on protection for the worksheet in $planSheet, checks all the allow boxes excel Insert and Delete columns and unlocks columns A-N
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',Justification='Does not change system state')]
    param (
        #The worksheet where protection is to be applied.
        [Parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet ,
        #Value of the "Protect Worksheet and Contents of locked cells" check box. Initially FALSE. use -IsProtected:$false to turn off it it has been switched on
        [switch]$IsProtected,
        #If provided sets all the ALLOW options to true or false and then allows them to be changed individually
        [switch]$AllowAll,
        #Opposite of the value in the 'Select locked cells' check box. Set to allow when Protect is first enabled
        [switch]$BlockSelectLockedCells,
        #Opposite of the value in the 'Select unlocked cells' check box. Set to allow when Protect is first enabled
        [switch]$BlockSelectUnlockedCells,
        #Value of the 'Format Cells' check box. Set to block when Protect is first enabled
        [switch]$AllowFormatCells,
        #Value of the 'Format Columns' check box. Set to block when Protect is first enabled
        [switch]$AllowFormatColumns,
        #Value of the 'Format Rows' check box. Set to block when Protect is first enabled
        [switch]$AllowFormatRows,
        #Value of the 'Insert Columns' check box. Set to block when Protect is first enabled
        [switch]$AllowInsertColumns,
        #Value of the 'Insert Columns' check box. Set to block when Protect is first enabled
        [switch]$AllowInsertRows,
        #Value of the 'Insert Hyperlinks' check box. Set to block when Protect is first enabled
        [switch]$AllowInsertHyperlinks,
        #Value of the 'Delete Columns' check box. Set to block when Protect is first enabled
        [switch]$AllowDeleteColumns,
        #Value of the 'Delete Rows' check box. Set to block when Protect is first enabled
        [switch]$AllowDeleteRows,
        #Value of the 'Sort' check box. Set to block when Protect is first enabled
        [switch]$AllowSort,
        #Value of the 'Use AutoFilter' check box. Set to block when Protect is first enabled
        [switch]$AllowAutoFilter,
        #Value of the 'Use PivotTable and PivotChart' check box. Set to block when Protect is first enabled
        [switch]$AllowPivotTables,
        ##Opposite of the value in the 'Edit objects' check box. Set to allow when Protect is first enabled
        [switch]$BlockEditObject,
        ##Opposite of the value in the 'Edit Scenarios' check box. Set to allow when Protect is first enabled
        [switch]$BlockEditScenarios,
        #Address range for cells to lock in the form "A:Z" or "1:10" or "A1:Z10"
        [string]$LockAddress,
         #Address range for cells to Unlock in the form "A:Z" or "1:10" or "A1:Z10"
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

    if ($UnlockAddress) {
        Set-ExcelRange     -Range $WorkSheet.cells[$UnlockAddress] -Locked:$false
    }
    if ($lockAddress) {
        Set-ExcelRange     -Range $WorkSheet.cells[$UnlockAddress] -Locked
    }
}