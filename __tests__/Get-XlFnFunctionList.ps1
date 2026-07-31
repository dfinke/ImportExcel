<#
    .SYNOPSIS
    Regenerates, from a real Excel installation, the list of functions Excel stores with an "_xlfn." prefix.

    .DESCRIPTION
    The table in Private\XlFnFormula.ps1 was generated with this script and can be re-verified or extended
    with it. It is not a Pester file and cannot run in CI: it drives the locally installed Excel through COM.

    Method: every candidate function is entered into a worksheet cell through the Excel object model (so
    Excel itself parses it), the workbook is saved, and the raw sheet XML is read back to see exactly what
    Excel stored - the function name alone, or the name carrying an "_xlfn." / "_xlfn._xlws." prefix. Excel
    validates argument shapes at entry, so several argument variants are tried, with hand-built formulas for
    functions whose arguments cannot be guessed (LET, LAMBDA and friends).

    The result is compared with the module's live table; any difference is reported. A candidate Excel
    refuses to enter at all (unknown to that Excel build) is reported as UNVERIFIED rather than guessed at.

    .EXAMPLE
    .\Get-XlFnFunctionList.ps1 | Format-Table
    Lists every candidate with the prefix Excel stored and whether the module's table agrees.

    .EXAMPLE
    .\Get-XlFnFunctionList.ps1 | Where-Object Verdict -ne OK
    Shows only the discrepancies - an empty result means the module's table matches this Excel build.
#>
[CmdletBinding()]
param(
    # Extra function names to test beyond the built-in candidate set - use when Excel gains new functions.
    [String[]]$AdditionalCandidates = @()
)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\..\ImportExcel.psd1" -Force

$candidates = @(
    #Post-2007 functions: expected to be stored with a prefix
    'AGGREGATE', 'BETA.DIST', 'BETA.INV', 'BINOM.DIST', 'BINOM.INV', 'CEILING.PRECISE', 'CHISQ.DIST',
    'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST', 'CONFIDENCE.NORM', 'CONFIDENCE.T',
    'COVARIANCE.P', 'COVARIANCE.S', 'ERF.PRECISE', 'ERFC.PRECISE', 'EXPON.DIST', 'F.DIST', 'F.DIST.RT',
    'F.INV', 'F.INV.RT', 'F.TEST', 'FLOOR.PRECISE', 'GAMMA.DIST', 'GAMMA.INV', 'GAMMALN.PRECISE',
    'HYPGEOM.DIST', 'LOGNORM.DIST', 'LOGNORM.INV', 'MODE.MULT', 'MODE.SNGL', 'NEGBINOM.DIST', 'NORM.DIST',
    'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'PERCENTILE.EXC', 'PERCENTILE.INC', 'PERCENTRANK.EXC',
    'PERCENTRANK.INC', 'POISSON.DIST', 'QUARTILE.EXC', 'QUARTILE.INC', 'RANK.AVG', 'RANK.EQ', 'STDEV.P',
    'STDEV.S', 'T.DIST', 'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST', 'VAR.P', 'VAR.S',
    'WEIBULL.DIST', 'Z.TEST',
    'ACOT', 'ACOTH', 'ARABIC', 'BASE', 'BINOM.DIST.RANGE', 'BITAND', 'BITLSHIFT', 'BITOR', 'BITRSHIFT',
    'BITXOR', 'CEILING.MATH', 'COMBINA', 'COT', 'COTH', 'CSC', 'CSCH', 'DAYS', 'DECIMAL', 'ENCODEURL',
    'FILTERXML', 'FLOOR.MATH', 'FORMULATEXT', 'GAMMA', 'GAUSS', 'IFNA', 'IMCOSH', 'IMCOT', 'IMCSC',
    'IMCSCH', 'IMSEC', 'IMSECH', 'IMSINH', 'IMTAN', 'ISFORMULA', 'ISOWEEKNUM', 'MUNIT', 'NUMBERVALUE',
    'PDURATION', 'PERMUTATIONA', 'PHI', 'RRI', 'SEC', 'SECH', 'SHEET', 'SHEETS', 'SKEW.P', 'UNICHAR',
    'UNICODE', 'WEBSERVICE', 'XOR',
    'FORECAST.ETS', 'FORECAST.ETS.CONFINT', 'FORECAST.ETS.SEASONALITY', 'FORECAST.ETS.STAT',
    'FORECAST.LINEAR',
    'CONCAT', 'IFS', 'MAXIFS', 'MINIFS', 'SWITCH', 'TEXTJOIN',
    'ARRAYTOTEXT', 'BYCOL', 'BYROW', 'CHOOSECOLS', 'CHOOSEROWS', 'DROP', 'EXPAND', 'FILTER', 'GROUPBY',
    'HSTACK', 'IMAGE', 'ISOMITTED', 'LAMBDA', 'LET', 'MAKEARRAY', 'MAP', 'PERCENTOF', 'PIVOTBY',
    'RANDARRAY', 'REDUCE', 'REGEXEXTRACT', 'REGEXREPLACE', 'REGEXTEST', 'SCAN', 'SEQUENCE', 'SORT',
    'SORTBY', 'STOCKHISTORY', 'TAKE', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTSPLIT', 'TOCOL', 'TOROW',
    'TRIMRANGE', 'UNIQUE', 'VALUETOTEXT', 'VSTACK', 'WRAPCOLS', 'WRAPROWS', 'XLOOKUP', 'XMATCH',
    #Functions Excel stores UNPREFIXED although some references claim otherwise
    'NETWORKDAYS.INTL', 'WORKDAY.INTL', 'ISO.CEILING',
    #Controls: pre-2007 functions which must never be prefixed
    'SUM', 'CONCATENATE', 'IFERROR', 'SUMIFS', 'VLOOKUP', 'TEXT', 'ZTEST'
) + $AdditionalCandidates

#Excel validates argument shapes when a formula is entered; these need specific arguments
$specialArgs = @{
    'LET'          = '=LET(x,1,x)'
    'LAMBDA'       = '=LAMBDA(x,x)(1)'
    'BYROW'        = '=BYROW(Z1:Z2,LAMBDA(r,SUM(r)))'
    'BYCOL'        = '=BYCOL(Z1:Z2,LAMBDA(c,SUM(c)))'
    'MAP'          = '=MAP(Z1:Z2,LAMBDA(v,v))'
    'REDUCE'       = '=REDUCE(0,Z1:Z2,LAMBDA(a,v,a+v))'
    'SCAN'         = '=SCAN(0,Z1:Z2,LAMBDA(a,v,a+v))'
    'MAKEARRAY'    = '=MAKEARRAY(2,2,LAMBDA(r,c,r*c))'
    'ISOMITTED'    = '=LAMBDA(x,ISOMITTED(x))(1)'
    'SWITCH'       = '=SWITCH(1,1,"a")'
    'GROUPBY'      = '=GROUPBY(Z1:Z2,Z1:Z2,LAMBDA(v,SUM(v)))'
    'PIVOTBY'      = '=PIVOTBY(Z1:Z2,Z1:Z2,Z1:Z2,LAMBDA(v,SUM(v)))'
    'HYPGEOM.DIST' = '=HYPGEOM.DIST(1,2,2,4,TRUE)'
    'RANK.AVG'     = '=RANK.AVG(1,Z1:Z2)'
    'RANK.EQ'      = '=RANK.EQ(1,Z1:Z2)'
    'Z.TEST'       = '=Z.TEST(Z1:Z2,0)'
    'ZTEST'        = '=ZTEST(Z1:Z2,0)'
    'MAXIFS'       = '=MAXIFS(Z1:Z2,Z1:Z2,1)'
    'MINIFS'       = '=MINIFS(Z1:Z2,Z1:Z2,1)'
    'SUMIFS'       = '=SUMIFS(Z1:Z2,Z1:Z2,1)'
}
$argVariants = '=NAME(1)', '=NAME(1,1)', '=NAME(1,1,1)', '=NAME("a")', '=NAME()', '=NAME(Z1:Z2)',
               '=NAME(1,1,1,1)', '=NAME("a","a")', '=NAME(Z1:Z2,1)'

$workFile = Join-Path ([IO.Path]::GetTempPath()) ("XlFnProbe{0:yyyyMMddHHmmss}.xlsx" -f (Get-Date))
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false
$excelApp.DisplayAlerts = $false
$results = [ordered]@{}
try {
    $workbook = $excelApp.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Range('Z1').Value2 = 1
    $sheet.Range('Z2').Value2 = 2
    Write-Verbose -Message ("Testing {0} candidate functions in Excel {1}" -f $candidates.Count, $excelApp.Version)
    $row = 0
    foreach ($name in $candidates) {
        $row ++
        $cell = $sheet.Cells.Item($row, 1)
        $entered = $false
        $tries = if ($specialArgs.Contains($name)) { @($specialArgs[$name]) }
                 else { foreach ($v in $argVariants) { $v -replace 'NAME', $name } }
        foreach ($try in $tries) {
            try { $cell.Formula2 = $try; $entered = $true; break } catch { }
        }
        $results[$name] = [pscustomobject][ordered]@{ Function = $name; Row = $row; Entered = $entered; ExcelStores = $null; ModuleTable = $null; Verdict = $null }
    }
    $workbook.SaveAs($workFile, 51)   #xlOpenXMLWorkbook
    $workbook.Close($false)
}
finally {
    $excelApp.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp)
}

#Read back what Excel stored for each row
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($workFile)
try {
    $entry = $zip.Entries | Where-Object { $_.FullName -match 'xl/worksheets/sheet1\.xml$' }
    $reader = [System.IO.StreamReader]::new($entry.Open())
    $sheetXml = [xml]$reader.ReadToEnd()
    $reader.Dispose()
    $nsManager = [System.Xml.XmlNamespaceManager]::new($sheetXml.NameTable)
    $nsManager.AddNamespace('d', $sheetXml.DocumentElement.NamespaceURI)
    #the module's live table, for comparison
    $moduleTable = & (Get-Module ImportExcel) { $script:XlFnFunctionPrefix }
    foreach ($result in $results.Values) {
        $result.ModuleTable = "$($moduleTable[$result.Function])"
        if (-not $result.Entered) {
            $result.ExcelStores = 'UNVERIFIED - Excel rejected every argument variant'
            $result.Verdict = 'UNVERIFIED'
            continue
        }
        $storedFormula = $sheetXml.SelectSingleNode("//d:c[@r='A$($result.Row)']/d:f", $nsManager).InnerText
        $result.ExcelStores = if ($storedFormula -match "(_xlfn\._xlws\.|_xlfn\.)$([regex]::Escape($result.Function))\s*\(") { $Matches[1] }
                              elseif ($storedFormula -match "(?<![\w.])$([regex]::Escape($result.Function))\s*\(")          { '' }
                              else { "UNRECOGNISED: $storedFormula" }
        $result.Verdict = if ($result.ExcelStores -eq $result.ModuleTable) { 'OK' } else { 'MISMATCH' }
    }
}
finally {
    $zip.Dispose()
    Remove-Item -Path $workFile -ErrorAction SilentlyContinue
}
$results.Values
