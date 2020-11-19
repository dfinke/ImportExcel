﻿#Requires -Modules Pester
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidAssignmentToAutomaticVariable", "", Justification='Sets IsWindows on pre-6.0 only')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','',Justification='False Positives')]
param()
if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}
if ($null -eq $IsWindows) {$IsWindows = [environment]::OSVersion.Platform -like "win*"}
$WarningAction = "SilentlyContinue"
Describe ExportExcel {
    . "$PSScriptRoot\Samples\Samples.ps1"
    if (Get-process -Name Excel,xlim -ErrorAction SilentlyContinue) {
        It "Excel is open" {
            $Warning = "You need to close Excel before running the tests."
            Write-Warning -Message $Warning
            Set-ItResult -Inconclusive -Because $Warning
        }
        return
    }

    Context "#Example 1      # Creates and opens a file with the right number of rows and columns" {
        $path = "TestDrive:\test.xlsx"
        Remove-item -Path $path -ErrorAction SilentlyContinue
        #Test with a maximum of 100 processes for speed; export all properties, then export smaller subsets.
        $processes = Get-Process | Where-Object {$_.StartTime} | Select-Object -First 100 -Property * -ExcludeProperty Parent
        $propertyNames = $Processes[0].psobject.properties.name
        $rowcount = $Processes.Count
        $Processes | Export-Excel $path  #-show

        it "Created a new file                                                                     " {
            Test-Path -Path $path -ErrorAction SilentlyContinue         | Should      -Be $true
        }

       # it "Started Excel to display the file                                                      " {
       #     Get-process -Name Excel, xlim -ErrorAction SilentlyContinue  | Should -Not -BeNullOrEmpty
       # }
       #Start-Sleep -Seconds 5 ;

        #Open-ExcelPackage with -Create is tested in Export-Excel
        #This is a test of  using it with -KillExcel
        #TODO Need to test opening pre-existing file with no -create switch (and graceful failure when file does not exist) somewhere else
        $Excel = Open-ExcelPackage -Path $path -KillExcel
        it "Killed Excel when Open-Excelpackage was told to                                        " {
            Get-process -Name Excel, xlim -ErrorAction SilentlyContinue  | Should      -BeNullOrEmpty
        }

        it "Created 1 worksheet, named 'Sheet1'                                                    " {
            $Excel.Workbook.Worksheets.count                            | Should      -Be 1
            $Excel.Workbook.Worksheets["Sheet1"]                        | Should -Not -BeNullOrEmpty
        }

        it "Added a 'Sheet1' property to the Package object                                        " {
            $Excel.Sheet1                                               | Should -Not -BeNullOrEmpty
        }

        $ws = $Excel.Workbook.Worksheets[1]
        it "Created the worksheet with the expected name, number of rows and number of columns     " {
            $ws.Name                                                    | Should      -Be "sheet1"
            $ws.Dimension.Columns                                       | Should      -Be  $propertyNames.Count
            $ws.Dimension.Rows                                          | Should      -Be ($rowcount + 1)
        }

        $headingNames = $ws.cells["1:1"].Value
        it "Created the worksheet with the correct header names                                    " {
            foreach ($p in $propertyNames) {
                $headingnames -contains $p                              | Should      -Be $true
            }
        }

        it "Formatted the process StartTime field as 'localized Date-Time'                         " {
            $STHeader = $ws.cells["1:1"].where( {$_.Value -eq "StartTime"})[0]
            $STCell = $STHeader.Address -replace '1$', '2'
            $ws.cells[$stcell].Style.Numberformat.NumFmtID              | Should      -Be 22
        }

        it "Formatted the process ID field as 'General'                                            " {
            $IDHeader = $ws.cells["1:1"].where( {$_.Value -eq "ID"})[0]
            $IDCell = $IDHeader.Address -replace '1$', '2'
            $ws.cells[$IDcell].Style.Numberformat.NumFmtID              | Should      -Be 0
        }
    }

    Context "                # NoAliasOrScriptPropeties -ExcludeProperty and -DisplayPropertySet work" {
        $path = "TestDrive:\test.xlsx"
        Remove-item -Path $path  -ErrorAction SilentlyContinue
        $processes = Get-Process | Select-Object -First 100
        $propertyNames = $Processes[0].psobject.properties.where( {$_.MemberType -eq 'Property'}).name
        $rowcount = $Processes.Count
        #Test -NoAliasOrScriptPropeties option and creating a range with a name which needs illegal chars removing - check this sends back a warning
        $warnVar = $null
        $Processes | Export-Excel $path -NoAliasOrScriptPropeties  -RangeName "No Spaces" -WarningVariable warnvar -WarningAction SilentlyContinue

        $Excel = Open-ExcelPackage -Path $path
        $ws = $Excel.Workbook.Worksheets[1]
        it "Created a new file with Alias & Script Properties removed.                             " {
            $ws.Name                                                    | Should      -Be "sheet1"
            $ws.Dimension.Columns                                       | Should      -Be  $propertyNames.Count
            $ws.Dimension.Rows                                          | Should      -Be  ($rowcount + 1 ) # +1 for the header.
        }
        it "Created a Range - even though the name given was invalid.                              " {
            $ws.Names["No_spaces"]                                      | Should -Not -BeNullOrEmpty
            $ws.Names["No_spaces"].End.Column                           | Should      -Be  $propertyNames.Count
            $ws.names["No_spaces"].End.Row                              | Should      -Be  ($rowcount + 1 ) # +1 for the header.
            $warnVar.Count                                              | Should      -Be  1
        }
        #This time use clearsheet instead of deleting the file test -Exclude properties, including wildcards.
        $Processes | Export-Excel $path -ClearSheet  -NoAliasOrScriptPropeties  -ExcludeProperty SafeHandle, threads, modules, MainModule, StartInfo, MachineName, MainWindow*, M*workingSet

        $Excel = Open-ExcelPackage -Path $path
        $ws = $Excel.Workbook.Worksheets[1]
        it "Created a new file with further properties excluded and cleared the old sheet          " {
            $ws.Name                                                    | Should      -Be "sheet1"
            $ws.Dimension.Columns                                       | Should      -Be ($propertyNames.Count - 10)
            $ws.Dimension.Rows                                          | Should      -Be ($rowcount + 1)  # +1 for the header
        }

        $propertyNames = $Processes[0].psStandardmembers.DefaultDisplayPropertySet.ReferencedPropertyNames
        Remove-item -Path $path -ErrorAction SilentlyContinue
        #Test -DisplayPropertySet
        $Processes | Export-Excel $path -DisplayPropertySet

        $Excel = Open-ExcelPackage -Path $path
        $ws = $Excel.Workbook.Worksheets[1]
        it "Created a new file with just the members of the Display Property Set                   " {
            $ws.Name                                                    | Should      -Be "sheet1"
            $ws.Dimension.Columns                                       | Should      -Be  $propertyNames.Count
            $ws.Dimension.Rows                                          | Should      -Be ($rowcount + 1)
        }
    }

    Context "#Example 2      # Exports a list of numbers and applies number format " {

        $path = "TestDrive:\test.xlsx"
        Remove-item -Path $path -ErrorAction SilentlyContinue
        #testing -ReturnRange switch and applying number format to Formulas as well as values.
        $returnedRange =   @($null, -1, 0, 34, 777, "", -0.5, 119, -0.1, 234, 788,"=A9+A10")   | Export-Excel -NumberFormat '[Blue]$#,##0.00;[Red]-$#,##0.00' -Path $path -ReturnRange
        it "Created a new file and returned the expected range                                     " {
            Test-Path -Path $path -ErrorAction SilentlyContinue         | Should      -Be $true
            $returnedRange                                              | Should      -Be "A1:A12"
        }

        $Excel = Open-ExcelPackage -Path $path
        it "Created 1 worksheet                                                                    " {
            $Excel.Workbook.Worksheets.count                            | Should      -Be 1
        }

        $ws = $Excel.Workbook.Worksheets[1]
        it "Created the worksheet with the expected name, number of rows and number of columns     " {
            $ws.Name                                                    | Should      -Be "sheet1"
            $ws.Dimension.Columns                                       | Should      -Be  1
            $ws.Dimension.End.Row                                       | Should      -Be  12
        }

        it "Set the default style for the sheet as expected                                        " {
            $ws.cells.Style.Numberformat.Format                         | Should      -Be  '[Blue]$#,##0.00;[Red]-$#,##0.00'
        }

        it "Set the default style and set values for Cells as expected, handling null,0 and ''     " {
            $ws.cells[1, 1].Style.Numberformat.Format                   | Should      -Be  '[Blue]$#,##0.00;[Red]-$#,##0.00'
            $ws.cells[1, 1].Value                                       | Should      -BeNullorEmpty
            $ws.cells[2, 1].Value                                       | Should      -Be -1
            $ws.cells[3, 1].Value                                       | Should      -Be 0
            $ws.cells[5, 1].Value                                       | Should      -Be 777
            $ws.cells[6, 1].Value                                       | Should      -Be ""
            $ws.cells[4, 1].Style.Numberformat.Format                   | Should      -Be  '[Blue]$#,##0.00;[Red]-$#,##0.00'

        }
    }

    Context "                # Number format parameter" {
        BeforeAll {
            $path = "TestDrive:\test.xlsx"
            Remove-Item -Path  $path -ErrorAction SilentlyContinue
            1..10  | Export-Excel -Path $path -Numberformat 'Number'
            1..10  | Export-Excel -Path $path -Numberformat 'Percentage' -Append
            21..30 | Export-Excel -Path $path -Numberformat 'Currency'   -StartColumn 3
            $excel = Open-ExcelPackage -Path   $path
            $ws = $excel.Workbook.Worksheets[1]
        }
        it "Set the worksheet default number format correctly                                      " {
            $ws.Cells.Style.Numberformat.Format                         | Should      -Be "0.00"
        }
        it "Set number formats on specific blocks of cells                                         " {
            $ws.Cells["A2" ].Style.Numberformat.Format                  | Should      -Be "0.00"
            $ws.Cells["c19"].Style.Numberformat.Format                  | Should      -Be "0.00"
            $ws.Cells["A20"].Style.Numberformat.Format                  | Should      -Be "0.00%"
            $ws.Cells["C6" ].Style.Numberformat.Format                  | Should      -Be (Expand-NumberFormat "currency")
        }
    }

    Context "#Examples 3 & 4 # Setting cells for different data types Also added test for URI type" {

        if ((Get-Culture).NumberFormat.CurrencySymbol -eq "£") {$OtherCurrencySymbol = "$"}
        else {$OtherCurrencySymbol = "£"}
        $path = "TestDrive:\test.xlsx"
        $warnVar = $null
        #Test correct export of different data types and number formats; test hyperlinks, test -NoNumberConversion test object is converted to a string with no warnings, test calcuation of formula
        Remove-item -Path $path -ErrorAction SilentlyContinue
        [PSCustOmobject][Ordered]@{
            Date             = Get-Date
            Formula1         = '=SUM(S2:T2)'
            String1          = 'My String'
            Float            = [math]::pi
            IPAddress        = '10.10.25.5'
            StrLeadZero      = '07670'
            StrComma         = '0,26'
            StrEngThousand   = '1,234.56'
            StrEuroThousand  = '1.555,83'
            StrDot           = '1.2'
            StrNegInt        = '-31'
            StrTrailingNeg   = '31-'
            StrParens        = '(123)'
            strLocalCurrency = ('{0}123{1}45' -f (Get-Culture).NumberFormat.CurrencySymbol,(Get-Culture).NumberFormat.CurrencyDecimalSeparator)
            strOtherCurrency = ('{0}123{1}45' -f $OtherCurrencySymbol ,(Get-Culture).NumberFormat.CurrencyDecimalSeparator)
            StrE164Phone     = '+32 (444) 444 4444'
            StrAltPhone1     = '+32 4 4444 444'
            StrAltPhone2     = '+3244444444'
            StrLeadSpace    = '  123'
            StrTrailSpace   = '123   '
            Link1            = [uri]"https://github.com/dfinke/ImportExcel"
            Link2            = "https://github.com/dfinke/ImportExcel"
            Link3            = "xl://internal/sheet1!A1"
            Link4            = "xl://internal/sheet1!C5"
            Link5            = (New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList "Sheet1!E2" , "Display Text")
            Process          = (Get-Process -Id $PID)
            TimeSpan         = [datetime]::Now.Subtract([datetime]::Today)
        } | Export-Excel  -NoNumberConversion IPAddress, StrLeadZero, StrAltPhone2  -Path $path -Calculate -WarningVariable $warnVar
        it "Created a new file                                                                     " {
            Test-Path -Path $path -ErrorAction SilentlyContinue         | Should      -Be $true
        }
        $Excel = Open-ExcelPackage -Path $path
        it "Created 1 worksheet with no warnings                                                   " {
            $Excel.Workbook.Worksheets.count                            | Should      -Be 1
            $warnVar                                                    | Should      -BeNullorEmpty
        }
        $ws = $Excel.Workbook.Worksheets[1]
        it "Created the worksheet with the expected name, number of rows and number of columns     " {
            $ws.Name                                                    | Should      -Be "sheet1"
            $ws.Dimension.Columns                                       | Should      -Be  27
            $ws.Dimension.Rows                                          | Should      -Be  2
        }
        it "Set a date     in Cell A2                                                              " {
            $ws.Cells[2, 1].Value.Gettype().name                        | Should      -Be  'DateTime'
        }
        it "Set a formula  in Cell B2                                                              " {
            $ws.Cells[2, 2].Formula                                     | Should      -Be  'SUM(S2:T2)'
        }
        it "Forced a successful calculation of the Value in Cell B2                                " {
            $ws.Cells[2, 2].Value                                       | Should      -Be  246
        }
        it "Set strings    in Cells E2, F2 and R2  (no number conversion)                          " {
            $ws.Cells[2,  5].Value.GetType().name                       | Should      -Be  'String'
            $ws.Cells[2,  6].Value.GetType().name                       | Should      -Be  'String'
            $ws.Cells[2, 18].Value.GetType().name                       | Should      -Be  'String'
        }
        it "Set numbers    in Cells K2,L2,M2   (diferent Negative integer formats)                 " {
            ($ws.Cells[2, 11].Value -is [valuetype] )                   | Should      -Be  $true
            ($ws.Cells[2, 12].Value -is [valuetype] )                   | Should      -Be  $true
            ($ws.Cells[2, 13].Value -is [valuetype] )                   | Should      -Be  $true
             $ws.Cells[2, 11].Value                                     | Should      -BeLessThan 0
             $ws.Cells[2, 12].Value                                     | Should      -BeLessThan 0
             $ws.Cells[2, 13].Value                                     | Should      -BeLessThan 0
        }
        it "Set external hyperlinks in Cells U2 and V2                                             " {
            $ws.Cells[2, 21].Hyperlink                                 | Should      -Be  "https://github.com/dfinke/ImportExcel"
            $ws.Cells[2, 22].Hyperlink                                 | Should      -Be  "https://github.com/dfinke/ImportExcel"
        }
        it "Set internal hyperlinks in Cells W2 and X2                                             " {
            $ws.Cells[2, 23].Hyperlink.Scheme                          | Should      -Be  "xl"
            $ws.Cells[2, 23].Hyperlink.ReferenceAddress                | Should      -Be  "sheet1!A1"
            $ws.Cells[2, 23].Hyperlink.Display                         | Should      -Be  "sheet1"
            $ws.Cells[2, 24].Hyperlink.Scheme                          | Should      -Be  "xl"
            $ws.Cells[2, 24].Hyperlink.ReferenceAddress                | Should      -Be  "sheet1!c5"
            $ws.Cells[2, 24].Hyperlink.Display                         | Should      -Be  "sheet1!c5"
            $ws.Cells[2, 25].Hyperlink.ReferenceAddress                | Should      -Be  "sheet1!E2"
            $ws.Cells[2, 25].Hyperlink.Display                         | Should      -Be  "Display Text"
        }
        it "Processed thousands according to local settings   (Cells H2 and I2)                    " {
            if ((Get-Culture).NumberFormat.NumberGroupSeparator -EQ ",") {
                ($ws.Cells[2, 8].Value -is [valuetype] )               | Should      -Be  $true
                 $ws.Cells[2, 9].Value.GetType().name                  | Should      -Be  'String'
            }
            elseif ((Get-Culture).NumberFormat.NumberGroupSeparator -EQ ".") {
                ($ws.Cells[2, 9].Value -is [valuetype] )               | Should      -Be  $true
                 $ws.Cells[2, 8].Value.GetType().name                  | Should      -Be  'String'
            }
        }
        it "Processed local currency as a number and other currency as a string (N2 & O2)          " {
            ($ws.Cells[2, 14].Value -is [valuetype] )                   | Should      -Be  $true
             $ws.Cells[2, 15].Value.GetType().name                      | Should      -Be  'String'
        }
        it "Processed numbers with spaces between digits as strings (P2 & Q2)                      " {
             $ws.Cells[2, 16].Value.GetType().name                      | Should      -Be  'String'
             $ws.Cells[2, 17].Value.GetType().name                      | Should      -Be  'String'
        }
        it "Processed numbers leading or trailing speaces as Numbers (S2 & T2)                     " {
            ($ws.Cells[2, 19].Value -is [valuetype] )                   | Should      -Be  $true
            ($ws.Cells[2, 20].Value -is [valuetype] )                   | Should      -Be  $true
        }
        it "Converted a nested object to a string (Y2)                                             " {
             $ws.Cells[2, 26].Value                                     | Should      -Match '^System\.Diagnostics\.Process\s+\(.*\)$'
        }
        it "Processed a timespan object (Z2)                                                       " {
             $ws.cells[2, 27].Value.ToOADate()                          | Should      -BeGreaterThan 0
             $ws.cells[2, 27].Value.ToOADate()                          | Should      -BeLessThan    1
             $ws.cells[2, 27].Style.Numberformat.Format                 | Should      -Be  '[h]:mm:ss'
        }
    }

    Context "#               # Setting cells for different data types with -noHeader" {

        $path = "TestDrive:\test.xlsx"
        Remove-item -Path $path -ErrorAction SilentlyContinue
        #Test -NoHeader & -NoNumberConversion
        [PSCustOmobject][Ordered]@{
            Date      = Get-Date
            Formula1  = '=SUM(F1:G1)'
            String1   = 'My String'
            String2   = 'a'
            IPAddress = '10.10.25.5'
            Number1   = '07670'
            Number2   = '0,26'
            Number3   = '1.555,83'
            Number4   = '1.2'
            Number5   = '-31'
            PhoneNr1  = '+32 44'
            PhoneNr2  = '+32 4 4444 444'
            PhoneNr3  = '+3244444444'
            Link      = [uri]"https://github.com/dfinke/ImportExcel"
        } | Export-Excel  -NoNumberConversion IPAddress, Number1  -Path $path -NoHeader
        it "Created a new file                                                                     " {
            Test-Path -Path $path -ErrorAction SilentlyContinue         | Should      -Be $true
        }

        $Excel = Open-ExcelPackage -Path $path
        it "Created 1 worksheet                                                                    " {
            $Excel.Workbook.Worksheets.count                            | Should      -Be 1
        }

        $ws = $Excel.Workbook.Worksheets[1]
        it "Created the worksheet with the expected name, number of rows and number of columns     " {
            $ws.Name                                                    | Should      -Be "sheet1"
            $ws.Dimension.Columns                                       | Should      -Be  14
            $ws.Dimension.Rows                                          | Should      -Be  1
        }

        it "Set a date      in Cell A1                                                             " {
            $ws.Cells[1, 1].Value.Gettype().name                        | Should      -Be  'DateTime'
        }

        it "Set a formula   in Cell B1                                                             " {
            $ws.Cells[1, 2].Formula                                     | Should      -Be  'SUM(F1:G1)'
        }

        it "Set strings     in Cells E1 and F1                                                     " {
            $ws.Cells[1, 5].Value.GetType().name                        | Should      -Be  'String'
            $ws.Cells[1, 6].Value.GetType().name                        | Should      -Be  'String'
        }

        it "Set a number    in Cell I1                                                             " {
            ($ws.Cells[1, 9].Value -is [valuetype] )                     | Should      -Be  $true
        }

        it "Set a hyperlink in Cell N1                                                             " {
            $ws.Cells[1, 14].Hyperlink                                   | Should      -Be  "https://github.com/dfinke/ImportExcel"
        }
    }

    Context "#Example 5      # Adding a single conditional format " {
        #Test  New-ConditionalText builds correctly
        $ct = New-ConditionalText -ConditionalType GreaterThan 525 -ConditionalTextColor ([System.Drawing.Color]::DarkRed) -BackgroundColor ([System.Drawing.Color]::LightPink)

        $path = "TestDrive:\test.xlsx"
        Remove-item -Path $path -ErrorAction SilentlyContinue
        #Test -ConditionalText with a single conditional spec.
        489, 668, 299, 777, 860, 151, 119, 497, 234, 788 | Export-Excel -Path $path -ConditionalText $ct

        it "Created a new file                                                                     " {
            Test-Path -Path $path -ErrorAction SilentlyContinue         | Should      -Be $true
        }

        #ToDo need to test applying conitional formatting to a pre-existing worksheet and removing = from formula
        $Excel = Open-ExcelPackage -Path $path
        $ws = $Excel.Workbook.Worksheets[1]

        it "Added one block of conditional formating for the data range                            " {
            $ws.ConditionalFormatting.Count                             | Should      -Be 1
            $ws.ConditionalFormatting[0].Address                        | Should      -Be ($ws.Dimension.Address)
        }

        $cf = $ws.ConditionalFormatting[0]
        it "Set the conditional formatting properties correctly                                    " {
            $cf.Formula                                                 | Should      -Be $ct.Text
            $cf.Type.ToString()                                         | Should      -Be $ct.ConditionalType
            #$cf.Style.Fill.BackgroundColor         | Should      -Be $ct.BackgroundColor
            # $cf.Style.Font.Color                   | Should -Be $ct.ConditionalTextColor  - have to compare r.g.b
        }
    }

    #Test adding mutliple conditional blocks and using the minimal syntax for New-ConditionalText
    $path = "TestDrive:\test.xlsx"
    Remove-item -Path $path -ErrorAction SilentlyContinue

    #Testing -Passthrough
    $Excel = Get-Service | Select-Object Name, Status, DisplayName, ServiceName |
        Export-Excel $path -PassThru  -ConditionalText $(
        New-ConditionalText Stop ([System.Drawing.Color]::DarkRed) ([System.Drawing.Color]::LightPink)
        New-ConditionalText Running ([System.Drawing.Color]::Blue) ([System.Drawing.Color]::Cyan)
    )
    $ws = $Excel.Workbook.Worksheets[1]

    Context "#Example 6      # Adding multiple conditional formats using short form syntax. " {
        it "Added two blocks of conditional formating for the data range                           " {
            $ws.ConditionalFormatting.Count                             | Should      -Be 2
            $ws.ConditionalFormatting[0].Address                        | Should      -Be ($ws.Dimension.Address)
            $ws.ConditionalFormatting[1].Address                        | Should      -Be ($ws.Dimension.Address)
        }
        it "Set the conditional formatting properties correctly                                    " {
            $ws.ConditionalFormatting[0].Text                           | Should      -Be "Stop"
            $ws.ConditionalFormatting[1].Text                           | Should      -Be "Running"
            $ws.ConditionalFormatting[0].Type                           | Should      -Be "ContainsText"
            $ws.ConditionalFormatting[1].Type                           | Should      -Be "ContainsText"
            #Add RGB Comparison
        }
    }
    Close-ExcelPackage -ExcelPackage $Excel

    Context "#Example 7      # Update-FirstObjectProperties works " {
        $Array = @()

        $Obj1 = [PSCustomObject]@{
            Member1 = 'First'
            Member2 = 'Second'
        }

        $Obj2 = [PSCustomObject]@{
            Member1 = 'First'
            Member2 = 'Second'
            Member3 = 'Third'
        }

        $Obj3 = [PSCustomObject]@{
            Member1 = 'First'
            Member2 = 'Second'
            Member3 = 'Third'
            Member4 = 'Fourth'
        }

        $Array = $Obj1, $Obj2, $Obj3
        #test Update-FirstObjectProperties
        $newarray = $Array | Update-FirstObjectProperties
        it "Outputs as many objects as it input                                                    " {
            $newarray.Count                                             | Should      -Be $Array.Count
        }
        it "Added properties to item 0                                                             " {
            $newarray[0].psobject.Properties.name.Count                 | Should      -Be 4
            $newarray[0].Member1                                        | Should      -Be 'First'
            $newarray[0].Member2                                        | Should      -Be 'Second'
            $newarray[0].Member3                                        | Should      -BeNullOrEmpty
            $newarray[0].Member4                                        | Should      -BeNullOrEmpty
        }
    }

    Context "#Examples 8 & 9 # Adding Pivot tables and charts from parameters" {
        $path = "TestDrive:\test.xlsx"
        #Test -passthru and -worksheetName creating a new, named, sheet in an existing file.
        $Excel = Get-Process |  Select-Object -first 20 -Property Name, cpu, pm, handles, company |  Export-Excel  $path -WorkSheetname Processes -PassThru
        #Testing -Excel Pacakage and adding a Pivot-table as a second step. Want to save and re-open it ...
        Export-Excel -ExcelPackage $Excel -WorkSheetname Processes -IncludePivotTable -PivotRows Company -PivotData PM -NoTotalsInPivot -PivotDataToColumn -Activate

        $Excel = Open-ExcelPackage  $path
        $PTws = $Excel.Workbook.Worksheets["ProcessesPivotTable"]
        $wCount = $Excel.Workbook.Worksheets.Count
        it "Added the named sheet and pivot table to the workbook                                  " {
            $excel.ProcessesPivotTable                                  | Should -Not -BeNullOrEmpty
            $PTws                                                       | Should -Not -BeNullOrEmpty
            $PTws.PivotTables.Count                                     | Should      -Be 1
            $Excel.Workbook.Worksheets["Processes"]                     | Should -Not -BeNullOrEmpty
            $Excel.Workbook.Worksheets.Count                            | Should      -BeGreaterThan 2
            $excel.Workbook.Worksheets["Processes"].Dimension.rows      | Should      -Be 21    #20 data + 1 header
        }
        it "Selected  the Pivottable page                                                          " {
            Set-ItResult -Pending -Because "Bug in EPPLus 4.5"
            $PTws.View.TabSelected                                      | Should      -Be $true
        }
        $pt = $PTws.PivotTables[0]
        it "Built the expected Pivot table                                                         " {
            $pt.RowFields.Count                                         | Should      -Be 1
            $pt.RowFields[0].Name                                       | Should      -Be "Company"
            $pt.DataFields.Count                                        | Should      -Be 1
            $pt.DataFields[0].Function                                  | Should      -Be "Count"
            $pt.DataFields[0].Field.Name                                | Should      -Be "PM"
            $PTws.Drawings.Count                                        | Should      -Be 0
        }
        #test adding pivot chart using the already open sheet
        $warnvar = $null
        Export-Excel -ExcelPackage $Excel -WorkSheetname Processes -IncludePivotTable -PivotRows Company -PivotData PM -IncludePivotChart -ChartType PieExploded3D -ShowCategory -ShowPercent  -NoLegend -WarningAction SilentlyContinue -WarningVariable warnvar
        $Excel = Open-ExcelPackage   $path
        it "Added a chart to the pivot table without rebuilding                                    " {
            $ws = $Excel.Workbook.Worksheets["ProcessesPivotTable"]
            $Excel.Workbook.Worksheets.Count                            | Should      -Be $wCount
            $ws.Drawings.count                                          | Should      -Be 1
            $ws.Drawings[0].ChartType.ToString()                        | Should      -Be "PieExploded3D"
        }
        it "Generated a message on re-processing the Pivot table                                   " {
            $warnVar                                                    | Should -Not -BeNullOrEmpty
        }
        #Test appending data extends pivot chart (with a warning) .
        $warnVar = $null
        Get-Process |  Select-Object -Last 20 -Property Name, cpu, pm, handles, company |   Export-Excel  $path -WorkSheetname Processes -Append -IncludePivotTable -PivotRows Company -PivotData PM -IncludePivotChart -ChartType PieExploded3D -WarningAction SilentlyContinue -WarningVariable warnvar
        $Excel = Open-ExcelPackage   $path
        $pt = $Excel.Workbook.Worksheets["ProcessesPivotTable"].PivotTables[0]
        it "Appended to the Worksheet and Extended the Pivot table                                 " {
            $Excel.Workbook.Worksheets.Count                            | Should      -Be $wCount
            $excel.Workbook.Worksheets["Processes"].Dimension.rows      | Should      -Be 41     #appended 20 rows to the previous total
            $pt.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref |
                 Should     be "A1:E41"
        }
        it "Generated a message on extending the Pivot table                                       " {
            $warnVar                                                    | Should -Not -BeNullOrEmpty
        }
    }

    Context "                # Add-Worksheet inserted sheets, moved them correctly, and copied a sheet" {
        $path = "TestDrive:\test.xlsx"
        #Test the -CopySource and -Movexxxx parameters for Add-Worksheet
        $Excel = Open-ExcelPackage  $path
        #At this point Sheets Should be in the order Sheet1, Processes, ProcessesPivotTable
        $null = Add-Worksheet -ExcelPackage $Excel -WorkSheetname "Processes" -MoveToEnd   # order now  Sheet1, ProcessesPivotTable, Processes
        $null = Add-Worksheet -ExcelPackage $Excel -WorkSheetname "NewSheet"  -MoveAfter "*" -CopySource ($excel.Workbook.Worksheets["Sheet1"]) # Now its NewSheet, Sheet1, ProcessesPivotTable, Processes
        $null = Add-Worksheet -ExcelPackage $Excel -WorkSheetname "Sheet1"    -MoveAfter "Processes"  # Now its NewSheet, ProcessesPivotTable, Processes, Sheet1
        $null = Add-Worksheet -ExcelPackage $Excel -WorkSheetname "Another"   -MoveToStart    # Now its Another, NewSheet, ProcessesPivotTable, Processes, Sheet1
        $null = Add-Worksheet -ExcelPackage $Excel -WorkSheetname "NearDone"  -MoveBefore 5   # Now its  Another, NewSheet, ProcessesPivotTable, Processes, NearDone ,Sheet1
        $null = Add-Worksheet -ExcelPackage $Excel -WorkSheetname "OneLast"   -MoveBefore "ProcessesPivotTable"   # Now its Another, NewSheet, Onelast, ProcessesPivotTable, Processes,NearDone ,Sheet1
        Close-ExcelPackage $Excel

        $Excel = Open-ExcelPackage  $path

        it "Got the Sheets in the right order                                                      " {
            $excel.Workbook.Worksheets[1].Name  | Should -Be "Another"
            $excel.Workbook.Worksheets[2].Name  | Should -Be "NewSheet"
            $excel.Workbook.Worksheets[3].Name  | Should -Be "Onelast"
            $excel.Workbook.Worksheets[4].Name  | Should -Be "ProcessesPivotTable"
            $excel.Workbook.Worksheets[5].Name  | Should -Be "Processes"
            $excel.Workbook.Worksheets[6].Name  | Should -Be "NearDone"
            $excel.Workbook.Worksheets[7].Name  | Should -Be "Sheet1"
        }

        it "Cloned 'Sheet1' to 'NewSheet'                                                          " {
            $newWs = $excel.Workbook.Worksheets["NewSheet"]
            $newWs.Dimension.Address                          | Should      -Be ($excel.Workbook.Worksheets["Sheet1"].Dimension.Address)
            $newWs.ConditionalFormatting.Count                | Should      -Be ($excel.Workbook.Worksheets["Sheet1"].ConditionalFormatting.Count)
            $newWs.ConditionalFormatting[0].Address.Address   | Should      -Be ($excel.Workbook.Worksheets["Sheet1"].ConditionalFormatting[0].Address.Address)
            $newWs.ConditionalFormatting[0].Formula           | Should      -Be ($excel.Workbook.Worksheets["Sheet1"].ConditionalFormatting[0].Formula)
        }

    }

    Context "                # Create and append with Start row and Start Column, inc ranges and Pivot table. " {
        $path = "TestDrive:\test.xlsx"
        remove-item -Path $path -ErrorAction SilentlyContinue
        #Catch warning
        $warnVar = $null
        #Test -Append with no existing sheet. Test adding a named pivot table from command line parameters and extending ranges when they're not specified explictly
        Get-Process | Select-Object -first 10 -Property Name, cpu, pm, handles, company  | Export-Excel -StartRow 3 -StartColumn 3 -BoldTopRow -IncludePivotTable  -PivotRows Company -PivotData PM -PivotTableName 'PTOffset' -Path $path -WorkSheetname withOffset -Append -PivotFilter Name -NoTotalsInPivot  -RangeName procs  -AutoFilter -AutoNameRange
        Get-Process | Select-Object -last  10 -Property Name, cpu, pm, handles, company  | Export-Excel -StartRow 3 -StartColumn 3 -BoldTopRow -IncludePivotTable  -PivotRows Company -PivotData PM -PivotTableName 'PTOffset' -Path $path -WorkSheetname withOffset -Append -WarningAction SilentlyContinue -WarningVariable warnvar
        $Excel = Open-ExcelPackage   $path
        $dataWs = $Excel.Workbook.Worksheets["withOffset"]
        $pt = $Excel.Workbook.Worksheets["PTOffset"].PivotTables[0]
        it "Created and appended to a sheet offset from the top left corner                        " {
            $dataWs.Cells[1, 1].Value                                   | Should      -BeNullOrEmpty
            $dataWs.Cells[2, 2].Value                                   | Should      -BeNullOrEmpty
            $dataWs.Cells[3, 3].Value                                   | Should -Not -BeNullOrEmpty
            $dataWs.Cells[3, 3].Style.Font.Bold                         | Should      -Be $true
            $dataWs.Dimension.End.Row                                   | Should      -Be 23
            $dataWs.names[0].Start.row                                  | Should      -Be 4   # StartRow + 1
            $dataWs.names[0].End.row                                    | Should      -Be $dataWs.Dimension.End.Row
            $dataWs.names[0].Name                                       | Should      -Be 'Name'
            $dataWs.names.Count                                         | Should      -Be 7    #  Name, cpu, pm, handles & company + Named Range "Procs" + xl one for autofilter
            $dataWs.cells[$dataws.Dimension].AutoFilter                 | Should      -Be true
            }
        it "Applied and auto-extended an autofilter                                                " {
            $dataWs.Names["_xlnm._FilterDatabase"].Start.Row            | Should      -Be 3  #offset
            $dataWs.Names["_xlnm._FilterDatabase"].Start.Column         | Should      -Be 3
            $dataWs.Names["_xlnm._FilterDatabase"].Rows                 | Should      -Be 21 #2 x 10 data + 1 header
            $dataWs.Names["_xlnm._FilterDatabase"].Columns              | Should      -Be 5  #Name, cpu, pm, handles & company
            $dataWs.Names["_xlnm._FilterDatabase"].AutoFilter           | Should      -Be $true
        }
        it "Created and auto-extended the named ranges                                             " {
            $dataWs.names["procs"].rows                                 | Should      -Be 21
            $dataWs.names["procs"].Columns                              | Should      -Be 5
            $dataWs.Names["CPU"].Rows                                   | Should      -Be 20
            $dataWs.Names["CPU"].Columns                                | Should      -Be 1
        }
        it "Created and extended the pivot table                                                   " {
            $pt.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref |
                                                                          Should     be "C3:G23"
            $pt.ColumGrandTotals                                        | Should      -Be $false
            $pt.RowGrandTotals                                          | Should      -Be $false
            $pt.Fields["Company"].IsRowField                            | Should      -Be $true
            $pt.Fields["PM"].IsDataField                                | Should      -Be $true
            $pt.Fields["Name"].IsPageField                              | Should      -Be $true
        }
        it "Generated a message on extending the Pivot table                                       " {
            $warnVar                                                    | Should -Not -BeNullOrEmpty
        }
    }

    Context "                # Create and append explicit and auto table and range extension" {
        $path = "TestDrive:\test.xlsx"
        #Test -Append automatically extends a table, even when it is not specified in the append command;
        Get-Process | Select-Object -first 10 -Property Name, cpu, pm, handles, company  | Export-Excel -Path $path  -TableName ProcTab -AutoNameRange   -WorkSheetname NoOffset -ClearSheet
        #Test number format applying to new data
        Get-Process | Select-Object -last  10 -Property Name, cpu, pm, handles, company  | Export-Excel -Path $path                     -AutoNameRange   -WorkSheetname NoOffset -Append -Numberformat 'Number'
        $Excel = Open-ExcelPackage   $path
        $dataWs = $Excel.Workbook.Worksheets["NoOffset"]
        #table should be 20 rows + header after extending the data. CPU range should be 1x20
        it "Created a new sheet and auto-extended a table and explicitly extended named ranges     " {
            $dataWs.Tables["ProcTab"].Address.Address                   | Should      -Be "A1:E21"
            $dataWs.Names["CPU"].Rows                                   | Should      -Be 20
            $dataWs.Names["CPU"].Columns                                | Should      -Be 1
        }
        it "Set the expected number formats                                                        " {
            $dataWs.cells["C2"].Style.Numberformat.Format               | Should      -Be "General"
            $dataWs.cells["C12"].Style.Numberformat.Format              | Should      -Be "0.00"
        }
        #Test extneding autofilter and range when explicitly specified in the append
        $excel = Get-Process | Select-Object -first 10 -Property Name, cpu, pm, handles, company  | Export-Excel -ExcelPackage $excel  -RangeName procs -AutoFilter   -WorkSheetname NoOffset -ClearSheet -PassThru
        Get-Process          | Select-Object -last  10 -Property Name, cpu, pm, handles, company  | Export-Excel -ExcelPackage $excel  -RangeName procs -AutoFilter   -WorkSheetname NoOffset -Append
        $Excel = Open-ExcelPackage   $path
        $dataWs = $Excel.Workbook.Worksheets["NoOffset"]

        it "Created a new sheet and explicitly extended named range and autofilter                 " {
            $dataWs.names["procs"].rows                                 | Should      -Be 21
            $dataWs.names["procs"].Columns                              | Should      -Be 5
            $dataWs.Names["_xlnm._FilterDatabase"].Rows                 | Should      -Be 21 #2 x 10 data + 1 header
            $dataWs.Names["_xlnm._FilterDatabase"].Columns              | Should      -Be 5  #Name, cpu, pm, handles & company
            $dataWs.Names["_xlnm._FilterDatabase"].AutoFilter           | Should      -Be $true
        }
    }

    Context "#Example 11     # Create and append with title, inc ranges and Pivot table" {
        $path = "TestDrive:\test.xlsx"
        #Test New-PivotTableDefinition builds definition using -Pivotfilter and -PivotTotals options.
        $ptDef = [ordered]@{}
        $ptDef += New-PivotTableDefinition -PivotTableName "PT1" -SourceWorkSheet 'Sheet1' -PivotRows "Status"  -PivotData @{'Status'  = 'Count'} -PivotTotals Columns -PivotFilter "StartType" -IncludePivotChart -ChartType BarClustered3D  -ChartTitle "Services by status" -ChartHeight 512 -ChartWidth 768 -ChartRow 10 -ChartColumn 0 -NoLegend -PivotColumns CanPauseAndContinue
        $ptDef += New-PivotTableDefinition -PivotTableName "PT2" -SourceWorkSheet 'Sheet2' -PivotRows "Company" -PivotData @{'Company' = 'Count'} -PivotTotalS Rows                             -IncludePivotChart -ChartType PieExploded3D -ShowPercent -WarningAction SilentlyContinue

        it "Built a pivot definition using New-PivotTableDefinition                                " {
            $ptDef.PT1.SourceWorkSheet                                  | Should -Be 'Sheet1'
            $ptDef.PT1.PivotRows                                        | Should -Be 'Status'
            $ptDef.PT1.PivotData.Status                                 | Should -Be 'Count'
            $ptDef.PT1.PivotFilter                                      | Should -Be 'StartType'
            $ptDef.PT1.IncludePivotChart                                | Should -Be  $true
            $ptDef.PT1.ChartType.tostring()                             | Should -Be 'BarClustered3D'
            $ptDef.PT1.PivotTotals                                      | Should -Be 'Columns'
        }
        Remove-Item -Path $path
        #Catch warning
        $warnvar = $null
        #Test create two data pages; as part of adding the second give both their own pivot table, test -autosize switch
        Get-Service | Select-Object    -Property Status, Name, DisplayName, StartType, CanPauseAndContinue | Export-Excel -Path $path  -AutoSize -TableName "All Services"  -TableStyle Medium1 -WarningVariable warnvar -WarningAction SilentlyContinue
        Get-Process | Select-Object    -Property Name, Company, Handles, CPU, VM      | Export-Excel -Path $path  -AutoSize -WorkSheetname 'sheet2' -TableName "Processes"     -TableStyle Light1 -Title "Processes" -TitleFillPattern Solid -TitleBackgroundColor ([System.Drawing.Color]::AliceBlue) -TitleBold -TitleSize 22 -PivotTableDefinition $ptDef
        $Excel = Open-ExcelPackage   $path
        $ws1 = $Excel.Workbook.Worksheets["Sheet1"]
        $ws2 = $Excel.Workbook.Worksheets["Sheet2"]

        if ($isWindows) {
          it "Set Column widths (with autosize)                                                      " {
            $ws1.Column(2).Width                                        | Should -Not -Be $ws1.DefaultColWidth
            $ws2.Column(1).width                                        | Should -Not -Be $ws2.DefaultColWidth
          }
        }

        it "Added tables to both sheets (handling illegal chars) and a title in sheet 2            " {
            $warnvar.count                                              | Should      -BeGreaterThan 0
            $ws1.tables.Count                                           | Should      -Be 1
            $ws2.tables.Count                                           | Should      -Be 1
            $ws1.Tables[0].Address.Start.Row                            | Should      -Be 1
            $ws2.Tables[0].Address.Start.Row                            | Should      -Be 2 #Title in row 1
            $ws1.Tables[0].Address.End.Address                          | Should      -Be $ws1.Dimension.End.Address
            $ws2.Tables[0].Address.End.Address                          | Should      -Be $ws2.Dimension.End.Address
            $ws2.Tables[0].Name                                         | Should      -Be "Processes"
            $ws2.Tables[0].StyleName                                    | Should      -Be "TableStyleLight1"
            $ws2.Cells["A1"].Value                                      | Should      -Be "Processes"
            $ws2.Cells["A1"].Style.Font.Bold                            | Should      -Be $true
            $ws2.Cells["A1"].Style.Font.Size                            | Should      -Be 22
            $ws2.Cells["A1"].Style.Fill.PatternType.tostring()          | Should      -Be "solid"
            $ws2.Cells["A1"].Style.Fill.BackgroundColor.Rgb             | Should      -Be "fff0f8ff"
        }

        $ptsheet1 = $Excel.Workbook.Worksheets["Pt1"]
        $ptsheet2 = $Excel.Workbook.Worksheets["Pt2"]
        $PT1 = $ptsheet1.PivotTables[0]
        $PT2 = $ptsheet2.PivotTables[0]
        $PC1 = $ptsheet1.Drawings[0]
        $PC2 = $ptsheet2.Drawings[0]
        it "Created the pivot tables linked to the right data.                                     " {
            $PT1.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.name|
                Should     be "All_services"
            $PT2.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.name |
                Should     be "Processes"
        }
        it "Set the other pivot tables and chart options from the definitions.                     " {
            $pt1.PageFields[0].Name                                     | Should      -Be 'StartType'
            $pt1.RowFields[0].Name                                      | Should      -Be 'Status'
            $pt1.DataFields[0].Field.name                               | Should      -Be 'Status'
            $pt1.DataFields[0].Function                                 | Should      -Be 'Count'
            $pt1.ColumGrandTotals                                       | Should      -Be $true
            $pt1.RowGrandTotals                                         | Should      -Be $false
            $pt2.ColumGrandTotals                                       | Should      -Be $false
            $pt2.RowGrandTotals                                         | Should      -Be $true
            $pc1.ChartType                                              | Should      -Be 'BarClustered3D'
            $pc1.From.Column                                            | Should      -Be 0                    #chart 1 at 0,10 chart 2 at 4,0 (default)
            $pc2.From.Column                                            | Should      -Be 4
            $pc1.From.Row                                               | Should      -Be 10
            $pc2.From.Row                                               | Should      -Be 0
            $pc1.Legend.Font                                            | Should      -BeNullOrEmpty           #Best check for legend removed.
            $pc2.Legend.Font                                            | Should -Not -BeNullOrEmpty
            $pc1.Title.Text                                             | Should      -Be 'Services by status'
            $pc2.DataLabel.ShowPercent                                  | Should      -Be $true
        }
    }

    Context "#Example 13     # Formatting and another way to do a pivot.  " {
        $path = "TestDrive:\test.xlsx"
        Remove-Item $path
        #Test freezing top row/first column, adding formats and a pivot table - from Add-Pivot table not a specification variable - after the export
        $excel = Get-Process | Select-Object -Property Name, Company, Handles, CPU, PM, NPM, WS | Export-Excel -Path $path -ClearSheet -WorkSheetname "Processes" -FreezeTopRowFirstColumn -PassThru
        $sheet = $excel.Workbook.Worksheets["Processes"]
        if ($isWindows) {$sheet.Column(1) | Set-ExcelRange -Bold -AutoFit }
        else            {$sheet.Column(1) | Set-ExcelRange -Bold  }
        $sheet.Column(2) | Set-ExcelRange -Width 29 -WrapText
        $sheet.Column(3) | Set-ExcelRange -HorizontalAlignment Right -NFormat "#,###"
        Set-ExcelRange -Address $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NFormat "#,###"
        Set-ExcelRange -Address $sheet.Column(4)  -HorizontalAlignment Right -NFormat "#,##0.0" -Bold
        Set-ExcelRange -Address $sheet.Row(1) -Bold -HorizontalAlignment Center
        Add-ConditionalFormatting -Worksheet $sheet -Range "D2:D1048576" -DataBarColor ([System.Drawing.Color]::Red)
        #test Add-ConditionalFormatting -passthru and using a range (and no worksheet)
        $rule = Add-ConditionalFormatting -passthru -Address $sheet.cells["C:C"] -RuleType TopPercent -ConditionValue 20 -Bold -StrikeThru
        Add-ConditionalFormatting -Worksheet $sheet -Range "G2:G1048576" -RuleType GreaterThan -ConditionValue "104857600" -ForeGroundColor ([System.Drawing.Color]::Red) -Bold -Italic -Underline -BackgroundColor  ([System.Drawing.Color]::Beige) -BackgroundPattern LightUp -PatternColor  ([System.Drawing.Color]::Gray)
        #Test Set-ExcelRange with a column
        if ($isWindows) { foreach ($c in 5..9) {Set-ExcelRange $sheet.Column($c)  -AutoFit } }
        Add-PivotTable -PivotTableName "PT_Procs" -ExcelPackage $excel -SourceWorkSheet 1 -PivotRows Company -PivotData  @{'Name' = 'Count'} -IncludePivotChart -ChartType ColumnClustered -NoLegend
        Export-Excel -ExcelPackage $excel -WorksheetName "Processes" -AutoNameRange #Test adding named ranges seperately from adding data.

        $excel = Open-ExcelPackage $path
        $sheet = $excel.Workbook.Worksheets["Processes"]
        it "Returned the rule when calling Add-ConditionalFormatting -passthru                     " {
            $rule                                                       | Should -Not -BeNullOrEmpty
            $rule.getType().fullname                                    | Should      -Be "OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingTopPercent"
            $rule.Style.Font.Strike                                     | Should -Be true
        }
        it "Applied the formating                                                                  " {
            $sheet                                                      | Should -Not -BeNullOrEmpty
            if ($isWindows) {
                $sheet.Column(1).width                                  | Should -Not -Be  $sheet.DefaultColWidth
                $sheet.Column(7).width                                  | Should -Not -Be  $sheet.DefaultColWidth
            }
            $sheet.Column(1).style.font.bold                            | Should      -Be  $true
            $sheet.Column(2).style.wraptext                             | Should      -Be  $true
            $sheet.Column(2).width                                      | Should      -Be  29
            $sheet.Column(3).style.horizontalalignment                  | Should      -Be  'right'
            $sheet.Column(4).style.horizontalalignment                  | Should      -Be  'right'
            $sheet.Cells["A1"].Style.HorizontalAlignment                | Should      -Be  'Center'
            $sheet.Cells['E2'].Style.HorizontalAlignment                | Should      -Be  'right'
            $sheet.Cells['A1'].Style.Font.Bold                          | Should      -Be  $true
            $sheet.Cells['D2'].Style.Font.Bold                          | Should      -Be  $true
            $sheet.Cells['E2'].style.numberformat.format                | Should      -Be  '#,###'
            $sheet.Column(3).style.numberformat.format                  | Should      -Be  '#,###'
            $sheet.Column(4).style.numberformat.format                  | Should      -Be  '#,##0.0'
            $sheet.ConditionalFormatting.Count                          | Should      -Be  3
            $sheet.ConditionalFormatting[0].type                        | Should      -Be  'Databar'
            $sheet.ConditionalFormatting[0].Color.name                  | Should      -Be  'ffff0000'
            $sheet.ConditionalFormatting[0].Address.Address             | Should      -Be  'D2:D1048576'
            $sheet.ConditionalFormatting[1].Style.Font.Strike           | Should      -Be  $true
            $sheet.ConditionalFormatting[1].type                        | Should      -Be  "TopPercent"
            $sheet.ConditionalFormatting[2].type                        | Should      -Be  'GreaterThan'
            $sheet.ConditionalFormatting[2].Formula                     | Should      -Be  '104857600'
            $sheet.ConditionalFormatting[2].Style.Font.Color.Color.Name | Should      -Be  'ffff0000'
        }
        it "Created the named ranges                                                               " {
            $sheet.Names.Count                                          | Should      -Be 7
            $sheet.Names[0].Start.Column                                | Should      -Be 1
            $sheet.Names[0].Start.Row                                   | Should      -Be 2
            $sheet.Names[0].End.Row                                     | Should      -Be $sheet.Dimension.End.Row
            $sheet.Names[0].Name                                        | Should      -Be $sheet.Cells['A1'].Value
            $sheet.Names[6].Start.Column                                | Should      -Be 7
            $sheet.Names[6].Start.Row                                   | Should      -Be 2
            $sheet.Names[6].End.Row                                     | Should      -Be $sheet.Dimension.End.Row
            $sheet.Names[6].Name                                        | Should      -Be $sheet.Cells['G1'].Value
        }
        it "Froze the panes                                                                        " {
            $sheet.view.Panes.Count                                     | Should      -Be 3
        }
        $ptsheet1 = $Excel.Workbook.Worksheets["Pt_procs"]

        it "Created the pivot table                                                                " {
            $ptsheet1                                                   | Should -Not -BeNullOrEmpty
            $ptsheet1.PivotTables[0].DataFields[0].Field.Name           | Should      -Be "Name"
            $ptsheet1.PivotTables[0].DataFields[0].Function             | Should      -Be "Count"
            $ptsheet1.PivotTables[0].RowFields[0].Name                  | Should      -Be "Company"
            $ptsheet1.PivotTables[0].CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref |
                Should     be $sheet.Dimension.address
        }
    }

    Context "                # Chart from MultiSeries.ps1 in the Examples\charts Directory" {
        $path = "TestDrive:\test.xlsx"
        Remove-Item -Path   $path -ErrorAction SilentlyContinue
        #Test we haven't missed any parameters on New-ChartDefinition which are on add chart or vice versa.

        $ParamChk1 =  (Get-command Add-ExcelChart          ).Parameters.Keys.where({-not (Get-command New-ExcelChartDefinition).Parameters.ContainsKey($_) }) | Sort-Object
        $ParamChk2 =  (Get-command New-ExcelChartDefinition).Parameters.Keys.where({-not (Get-command Add-ExcelChart          ).Parameters.ContainsKey($_) })
        it "Found the same parameters for Add-ExcelChart and New-ExcelChartDefinintion             " {
            $ParamChk1.count                                            | Should      -Be 3
            $ParamChk1[0]                                               | Should      -Be "PassThru"
            $ParamChk1[1]                                               | Should      -Be "PivotTable"
            $ParamChk1[2]                                               | Should      -Be "Worksheet"
            $ParamChk2.count                                            | Should      -Be 1
            $ParamChk2[0]                                               | Should      -Be "Header"
        }
        #Test Invoke-Sum
        $data = Invoke-Sum (Get-Process) Company Handles, PM, VirtualMemorySize
        it "Used Invoke-Sum to create a data set                                                   " {
            $data                                                       | Should -Not -BeNullOrEmpty
            $data.count                                                 | Should      -BeGreaterThan 1
            $data[1].Name                                               | Should -Not -BeNullOrEmpty
            $data[1].Handles                                            | Should -Not -BeNullOrEmpty
            $data[1].PM                                                 | Should -Not -BeNullOrEmpty
            $data[1].VirtualMemorySize                                  | Should -Not -BeNullOrEmpty
        }
        $c = New-ExcelChartDefinition -Title Stats -ChartType LineMarkersStacked   -XRange "Processes[Name]" -YRange "Processes[PM]", "Processes[VirtualMemorySize]" -SeriesHeader 'PM', 'VMSize'

        it "Created the Excel chart definition                                                     " {
            $c                                                          | Should -Not -BeNullOrEmpty
            $c.ChartType.gettype().name                                 | Should      -Be "eChartType"
            $c.ChartType.tostring()                                     | Should      -Be "LineMarkersStacked"
            $c.yrange -is [array]                                       | Should      -Be $true
            $c.yrange.count                                             | Should      -Be 2
            $c.yrange[0]                                                | Should      -Be "Processes[PM]"
            $c.yrange[1]                                                | Should      -Be "Processes[VirtualMemorySize]"
            $c.xrange                                                   | Should      -Be "Processes[Name]"
            $c.Title                                                    | Should      -Be "Stats"
            $c.Nolegend                                                 | Should -Not -Be $true
            $c.ShowCategory                                             | Should -Not -Be $true
            $c.ShowPercent                                              | Should -Not -Be $true
        }
        #Test creating a chart using -ExcelChartDefinition.
        $data | Export-Excel $path -AutoSize -TableName Processes -ExcelChartDefinition $c
        $excel = Open-ExcelPackage -Path $path
        $drawings = $excel.Workbook.Worksheets[1].drawings
        it "Used the Excel chart definition with Export-Excel                                      " {
            $drawings.count                                             | Should      -Be 1
            $drawings[0].ChartType                                      | Should      -Be "LineMarkersStacked"
            $drawings[0].Series.count                                   | Should      -Be 2
            $drawings[0].Series[0].Series                               | Should      -Be "'Sheet1'!Processes[PM]"
            $drawings[0].Series[0].XSeries                              | Should      -Be "'Sheet1'!Processes[Name]"
            $drawings[0].Series[1].Series                               | Should      -Be "'Sheet1'!Processes[VirtualMemorySize]"
            $drawings[0].Series[1].XSeries                              | Should      -Be "'Sheet1'!Processes[Name]"
            $drawings[0].Title.text                                     | Should      -Be "Stats"
        }
        Close-ExcelPackage $excel
    }

    Context "                # variation of plot.ps1 from Examples Directory using Add chart outside ExportExcel" {
        $path = "TestDrive:\test.xlsx"
        #Test inserting a fomual
        $excel = 0..360 | ForEach-Object {[pscustomobject][ordered]@{x = $_; Sinx = "=Sin(Radians(x)) "}} | Export-Excel -AutoNameRange -Path $path -WorkSheetname SinX -ClearSheet -FreezeFirstColumn -PassThru
        #Test-Add Excel Chart to existing data. Test add Conditional formatting with a formula
        Add-ExcelChart -Worksheet $excel.Workbook.Worksheets["Sinx"] -ChartType line -XRange "X" -YRange "Sinx" -SeriesHeader "Sin(x)" -Title "Graph of Sine X" -TitleBold -TitleSize 14 `
                       -Column 2 -ColumnOffSetPixels 35 -Width 800 -XAxisTitleText "Degrees" -XAxisTitleBold -XAxisTitleSize 12 -XMajorUnit 30 -XMinorUnit 10 -XMinValue 0 -XMaxValue 361  -XAxisNumberformat "000" `
                       -YMinValue -1.25 -YMaxValue 1.25 -YMajorUnit 0.25 -YAxisNumberformat "0.00" -YAxisTitleText "Sine" -YAxisTitleBold -YAxisTitleSize 12 `
                       -LegendSize 8 -legendBold  -LegendPosition Bottom
        Add-ConditionalFormatting -Worksheet $excel.Workbook.Worksheets["Sinx"] -Range "B2:B362" -RuleType LessThan -ConditionValue "=B1" -ForeGroundColor ([System.Drawing.Color]::Red)
        $ws = $Excel.Workbook.Worksheets["Sinx"]
        $d  = $ws.Drawings[0]
        It "Controled the axes and title and legend of the chart                                   " {
            $d.XAxis.MaxValue                                           | Should      -Be 361
            $d.XAxis.MajorUnit                                          | Should      -Be 30
            $d.XAxis.MinorUnit                                          | Should      -Be 10
            $d.XAxis.Title.Text                                         | Should      -Be "degrees"
            $d.XAxis.Title.Font.bold                                    | Should      -Be $true
            $d.XAxis.Title.Font.Size                                    | Should      -Be 12
            $d.XAxis.MajorUnit                                          | Should      -Be 30
            $d.XAxis.MinorUnit                                          | Should      -Be 10
            $d.XAxis.MinValue                                           | Should      -Be 0
            $d.XAxis.MaxValue                                           | Should      -Be 361
            $d.YAxis.Format                                             | Should      -Be "0.00"
            $d.Title.Text                                               | Should      -Be "Graph of Sine X"
            $d.Title.Font.Bold                                          | Should      -Be $true
            $d.Title.Font.Size                                          | Should      -Be 14
            $d.yAxis.MajorUnit                                          | Should      -Be 0.25
            $d.yAxis.MaxValue                                           | Should      -Be 1.25
            $d.yaxis.MinValue                                           | Should      -Be -1.25
            $d.Legend.Position.ToString()                               | Should      -Be "Bottom"
            $d.Legend.Font.Bold                                         | Should      -Be $true
            $d.Legend.Font.Size                                         | Should      -Be 8
            $d.ChartType.tostring()                                     | Should      -Be "line"
            $d.From.Column                                              | Should      -Be 2
        }
        It "Appplied conditional formatting to the data                                            " {
            $ws.ConditionalFormatting[0].Formula                        | Should      -Be "B1"
        }
        Close-ExcelPackage -ExcelPackage $excel -nosave
    }

    Context "                # Quick line chart" {
        $path = "TestDrive:\test.xlsx"
        Remove-Item -Path $path -ErrorAction SilentlyContinue
        #test drawing a chart when data doesn't have a string
        0..360 | ForEach-Object {[pscustomobject][ordered]@{x = $_; Sinx = "=Sin(Radians(x)) "}} | Export-Excel -AutoNameRange  -Path $path -LineChart
        $excel = Open-ExcelPackage -Path $path
        $ws = $excel.Sheet1
        $d = $ws.Drawings[0]
        it "Created the chart                                                                      " {
            $d.Title.text                                                 | Should      -BeNullOrEmpty
            $d.ChartType                                                  | Should      -Be "line"
            $d.Series[0].Header                                           | Should      -Be "Sinx"
            $d.Series[0].xSeries                                          | Should      -Be "'Sheet1'!A2:A362"
            $d.Series[0].Series                                           | Should      -Be "'Sheet1'!B2:B362"
        }

    }

    Context "                # Quick Pie chart and three icon conditional formating" {
        $path = "TestDrive:\Pie.xlsx"
        Remove-Item -Path $path -ErrorAction SilentlyContinue
        $range = Get-Process| Group-Object -Property company | Where-Object -Property name |
             Select-Object -Property Name, @{n="TotalPm";e={($_.group | Measure-Object -sum -Property pm).sum }} |
                 Export-Excel -NoHeader -AutoNameRange -path $path -ReturnRange  -PieChart -ShowPercent
        $Cf = New-ConditionalFormattingIconSet -Range ($range -replace "^.*:","B2:") -ConditionalFormat ThreeIconSet -Reverse -IconType Flags
        $ct = New-ConditionalText -Text "Microsoft" -ConditionalTextColor ([System.Drawing.Color]::Red) -BackgroundColor([System.Drawing.Color]::AliceBlue) -ConditionalType ContainsText
        it "Created the Conditional formatting rules                                               " {
            $cf.Formatter                                               | Should      -Be "ThreeIconSet"
            $cf.IconType                                                | Should      -Be "Flags"
            $cf.Range                                                   | Should      -Be ($range -replace "^.*:","B2:")
            $cf.Reverse                                                 | Should      -Be $true
            $ct.BackgroundColor.Name                                    | Should      -Be "AliceBlue"
            $ct.ConditionalTextColor.Name                               | Should      -Be "Red"
            $ct.ConditionalType                                         | Should      -Be "ContainsText"
            $ct.Text                                                    | Should      -Be "Microsoft"
        }
        #Test -ConditionalFormat & -ConditionalText
        Export-Excel -Path $path -ConditionalFormat $cf -ConditionalText $ct
        $excel = Open-ExcelPackage -Path $path
        $rows  = $range -replace "^.*?(\d+)$", '$1'
        $chart = $excel.Workbook.Worksheets["sheet1"].Drawings[0]
        $cFmt  = $excel.Workbook.Worksheets["sheet1"].ConditionalFormatting
        it "Created the chart with the right series                                                " {
            $chart.ChartType                                            | Should      -Be "PieExploded3D"
            $chart.series.series                                        | Should      -Be "'Sheet1'!B1:B$rows" #would be B2 and A2 if we had a header.
            $chart.series.Xseries                                       | Should      -Be "'Sheet1'!A1:A$rows"
            $chart.DataLabel.ShowPercent                                | Should      -Be $true
        }
        it "Created two Conditional formatting rules                                               " {
            $cFmt.Count                                                 | Should      -Be $true
            $cFmt.Where({$_.type -eq "ContainsText"})                   | Should -Not -BeNullOrEmpty
            $cFmt.Where({$_.type -eq "ThreeIconSet"})                   | Should -Not -BeNullOrEmpty
        }
    }

    Context "                # Awkward multiple tables" {
        $path = "TestDrive:\test.xlsx"
        #Test creating 3 on overlapping tables on the same page. Create rightmost the left most then middle.
        remove-item -Path $path -ErrorAction SilentlyContinue
        if ($IsLinux -or $IsMacOS) {
            $SystemFolder = '/etc'
        }
        else {
            $SystemFolder = 'C:\WINDOWS\system32'
        }
        $r = Get-ChildItem -path $SystemFolder -File

        "Biggest files" | Export-Excel -Path $path -StartRow 1 -StartColumn 7
        $r | Sort-Object length -Descending | Select-Object -First 14 Name, @{n="Size";e={$_.Length}}  |
            Export-Excel -Path $path -TableName FileSize -StartRow 2 -StartColumn 7 -TableStyle Medium2

        $r.extension | Group-Object | Sort-Object -Property count -Descending | Select-Object -First 12 Name, Count   |
            Export-Excel -Path $path -TableName ExtSize -Title "Frequent Extensions"  -TitleSize 11 -BoldTopRow

        $r | Group-Object -Property extension | Select-Object Name, @{n="Size"; e={($_.group  | Measure-Object -property length -sum).sum}} |
          Sort-Object -Property size -Descending | Select-Object -First 10 |
            Export-Excel -Path $path -TableName ExtCount -Title "Biggest extensions"  -TitleSize 11 -StartColumn 4 -AutoSize

        $excel = Open-ExcelPackage -Path $path
        $ws = $excel.Workbook.Worksheets[1]
        it "Created 3 tables                                                                       " {
            $ws.tables.count | Should -Be 3
        }
        it "Created the FileSize table in the right place with the right size and style            " {
            $ws.Tables["FileSize"].Address.Address                      | Should      -Be "G2:H16" #Insert at row 2, Column 7, 14 rows x 2 columns of data
            $ws.Tables["FileSize"].StyleName                            | Should      -Be "TableStyleMedium2"
        }
        it "Created the ExtSize  table in the right place with the right size and style            " {
            $ws.Tables["ExtSize"].Address.Address                      | should      -be "A2:B14" #tile, then 12 rows x 2 columns of data
            $ws.Tables["ExtSize"].StyleName                            | should      -be "TableStyleMedium6"
        }
        it "Created the ExtCount table in the right place with the right size                      " {
            $ws.Tables["ExtCount"].Address.Address                      | Should      -Be "D2:E12" #title, then 10 rows x 2 columns of data
        }
    }

    Context "                # Parameters and ParameterSets" {
        $Path = Join-Path (Resolve-Path 'TestDrive:').ProviderPath "test.xlsx"
        Remove-Item -Path $Path -ErrorAction SilentlyContinue
        $Processes = Get-Process | Select-Object -first 10 -Property Name, cpu, pm, handles, company

        it "Allows the default parameter set with Path".PadRight(87) {
            $ExcelPackage = $Processes | Export-Excel -Path $Path -PassThru
            $Worksheet = $ExcelPackage.Workbook.Worksheets[1]

            $ExcelPackage.File | Should -Be $Path
            $Worksheet.Cells['A1'].Value | Should -Be 'Name'
            $Worksheet.Tables | Should -BeNullOrEmpty
            $Worksheet.AutoFilterAddress | Should -BeNullOrEmpty
        }
        it "throws when the ExcelPackage is specified with either -path or -Now".PadRight(87) {
            $ExcelPackage = Export-Excel -Path $Path -PassThru
            {Export-Excel -ExcelPackage $ExcelPackage -Path $Path} | Should  -Throw 'Parameter set cannot be resolved using the specified named parameters'
            {Export-Excel -ExcelPackage $ExcelPackage -Now} | Should  -Throw 'Parameter set cannot be resolved using the specified named parameters'

            $Processes | Export-Excel -ExcelPackage $ExcelPackage
            Remove-Item -Path $Path
        }
        it "If TableName and AutoFilter provided AutoFilter will be ignored".PadRight(87) {
            $ExcelPackage = Export-Excel -Path $Path -PassThru -TableName 'Data' -AutoFilter
            $Worksheet = $ExcelPackage.Workbook.Worksheets[1]

            $Worksheet.Tables[0].Name | Should -Be 'Data'
            $Worksheet.AutoFilterAddress | Should -BeNullOrEmpty
        }
        it "Default Set with Path and TableName with generated name".PadRight(87) {
            $ExcelPackage = $Processes | Export-Excel -Path $Path -PassThru -TableName ''
            $Worksheet = $ExcelPackage.Workbook.Worksheets[1]

            $ExcelPackage.File | Should -Be $Path
            $Worksheet.Tables[0].Name | Should -Be 'Table1'
        }
        it "Now will use temp Path, set TableName with generated name and AutoSize".PadRight(87) {
            $ExcelPackage = $Processes | Export-Excel -Now -PassThru
            $Worksheet = $ExcelPackage.Workbook.Worksheets[1]

            $ExcelPackage.File.FullName   | Should -BeLike ([IO.Path]::GetTempPath() + '*')
            $Worksheet.Tables[0].Name      | Should -Be 'Table1'
            $Worksheet.AutoFilterAddress  | Should -BeNullOrEmpty
            if ($isWindows) {
                $Worksheet.Column(5).Width | Should -BeGreaterThan 9.5
            }
        }
        it "Now allows override of Path and TableName".PadRight(87) {
            $ExcelPackage = $Processes | Export-Excel -Now -PassThru -Path $Path -TableName:$false
            $Worksheet = $ExcelPackage.Workbook.Worksheets[1]

            $ExcelPackage.File | Should -Be $Path
            $Worksheet.Tables | Should -BeNullOrEmpty
            $Worksheet.AutoFilterAddress | Should -BeNullOrEmpty
            if ($isWindows) {
                $Worksheet.Column(5).Width | Should -BeGreaterThan 9.5
            }
        }
        <# Mock looks unreliable need to check
        Mock -CommandName 'Invoke-Item'
        it "Now will Show".PadRight(87) {
            $Processes | Export-Excel
            Assert-MockCalled -CommandName 'Invoke-Item' -Times 1 -Exactly -Scope 'It'
        }
        it "Now allows override of Show".PadRight(87) {
            $Processes | Export-Excel -Show:$false
            Assert-MockCalled -CommandName 'Invoke-Item' -Times 0 -Exactly -Scope 'It'
        }
        #>
        it "Now allows override of AutoSize and TableName to AutoFilter".PadRight(87) {
            $ExcelPackage = $Processes | Export-Excel -Now -PassThru -AutoSize:$false -AutoFilter
            $Worksheet = $ExcelPackage.Workbook.Worksheets[1]

            $Worksheet.Tables | Should -BeNullOrEmpty
            $Worksheet.AutoFilterAddress | Should -Not -BeNullOrEmpty
            [math]::Round($Worksheet.Column(5).Width, 2) | Should -Be 9.14
        }
        it "Now allows to set TableName".PadRight(87) {
            $ExcelPackage = $Processes | Export-Excel -Now -PassThru -TableName 'Data'
            $Worksheet = $ExcelPackage.Workbook.Worksheets[1]

            $Worksheet.Tables[0].Name | Should -Be 'Data'
            $Worksheet.AutoFilterAddress | Should -BeNullOrEmpty
            if ($isWindows) {
                $Worksheet.Column(5).Width | Should -BeGreaterThan 9.5
            }
        }
    }
}
