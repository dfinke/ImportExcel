﻿#Requires -Modules Pester

# $here = Split-Path -Parent $MyInvocation.MyCommand.Path
# Import-Module $here -Force -Verbose
Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force

if (Get-process -Name Excel,xlim -ErrorAction SilentlyContinue) {    Write-Warning -Message "You need to close Excel before running the tests." ; return}
Describe ExportExcel {

    Context "#Example 1      # Creates and opens a file with the right number of rows and columns" {
        $path = "$env:TEMP\Test.xlsx"
        Remove-item -Path $path -ErrorAction SilentlyContinue
        $processes = Get-Process
        $propertyNames = $Processes[0].psobject.properties.name
        $rowcount = $Processes.Count
        $Processes | Export-Excel $path  #-show

        it "Created a new file                                                                     " {
            Test-Path -Path $path -ErrorAction SilentlyContinue         | Should     be $true
        }

       # it "Started Excel to display the file                                                      " {
       #     Get-process -Name Excel, xlim -ErrorAction SilentlyContinue  | Should not benullorempty
       # }
       #Start-Sleep -Seconds 5 ;

        #Open-ExcelPackage with -Create is tested in Export-Excel
        #This is a test of  using it with -KillExcel
        #TODO Need to test opening pre-existing file with no -create switch (and graceful failure when file does not exist) somewhere else
        $Excel = Open-ExcelPackage -Path $path -KillExcel
        it "Killed Excel when Open-Excelpackage was told to                                        " {
            Get-process -Name Excel, xlim -ErrorAction SilentlyContinue  | Should     benullorempty
        }

        it "Created 1 worksheet                                                                    " {
            $Excel.Workbook.Worksheets.count                            | Should     be 1
        }

        $ws = $Excel.Workbook.Worksheets[1]
        it "Created the worksheet with the expected name, number of rows and number of columns     " {
            $ws.Name                                                    | Should     be "sheet1"
            $ws.Dimension.Columns                                       | Should     be  $propertyNames.Count
            $ws.Dimension.Rows                                          | Should     be ($rowcount + 1)
        }

        $headingNames = $ws.cells["1:1"].Value
        it "Created the worksheet with the correct header names                                    " {
            foreach ($p in $propertyNames) {
                $headingnames -contains $p                              | Should     be $true
            }
        }

        it "Formatted the process StartTime field as 'localized Date-Time'                          " {
            $STHeader = $ws.cells["1:1"].where( {$_.Value -eq "StartTime"})[0]
            $STCell = $STHeader.Address -replace '1$', '2'
            $ws.cells[$stcell].Style.Numberformat.NumFmtID              | Should     be 22
        }

        it "Formatted the process ID field as 'General'                                            " {
            $IDHeader = $ws.cells["1:1"].where( {$_.Value -eq "ID"})[0]
            $IDCell = $IDHeader.Address -replace '1$', '2'
            $ws.cells[$IDcell].Style.Numberformat.NumFmtID              | Should     be 0
        }
    }

    Context "                # NoAliasOrScriptProperties -ExcludeProperty and -DisplayPropertySet work" {
        $path = "$env:TEMP\Test.xlsx"
        Remove-item -Path $path  -ErrorAction SilentlyContinue
        $processes = Get-Process
        $propertyNames = $Processes[0].psobject.properties.where( {$_.MemberType -ne 'ScriptProperty' -and $_.MemberType -ne 'AliasProperty'}).name
        $rowcount = $Processes.Count
        #TestCreating a range with a name which needs illegal chars removing
        $warnVar = $null
        $Processes | Export-Excel $path -NoAliasOrScriptProperties  -RangeName "No Spaces" -WarningVariable warnvar -WarningAction SilentlyContinue

        $Excel = Open-ExcelPackage -Path $path
        $ws = $Excel.Workbook.Worksheets[1]
        it "Created a new file with alias & Script Properties removed.                             " {
            $ws.Name                                                    | Should     be "sheet1"
            $ws.Dimension.Columns                                       | Should     be  $propertyNames.Count
            $ws.Dimension.Rows                                          | Should     be  ($rowcount + 1 ) # +1 for the header.
        }
        it "Created a Range - even though the name given was invalid.                              " {
            $ws.Names["No_spaces"]                                      | Should not beNullOrEmpty
            $ws.Names["No_spaces"].End.Column                           | Should     be  $propertyNames.Count
            $ws.names["No_spaces"].End.Row                              | Should     be  ($rowcount + 1 ) # +1 for the header.
            $warnVar.Count                                              | Should     be  1
        }
        #This time use clearsheet instead of deleting the file
        $Processes | Export-Excel $path -NoAliasOrScriptProperties -ExcludeProperty SafeHandle, modules, MainModule, StartTime, Threads -ClearSheet

        $Excel = Open-ExcelPackage -Path $path
        $ws = $Excel.Workbook.Worksheets[1]
        it "Created a new file with a further 5 properties excluded and cleared the old sheet      " {
            $ws.Name                                                    | Should     be "sheet1"
            $ws.Dimension.Columns                                       | Should     be ($propertyNames.Count - 5)
            $ws.Dimension.Rows                                          | Should     be ($rowcount + 1)  # +1 for the header
        }

        $propertyNames = $Processes[0].psStandardmembers.DefaultDisplayPropertySet.ReferencedPropertyNames
        Remove-item -Path $path -ErrorAction SilentlyContinue
        $Processes | Export-Excel $path -DisplayPropertySet

        $Excel = Open-ExcelPackage -Path $path
        $ws = $Excel.Workbook.Worksheets[1]
        it "Created a new file with just the members of the Display Property Set                   " {
            $ws.Name                                                    | Should     be "sheet1"
            $ws.Dimension.Columns                                       | Should     be  $propertyNames.Count
            $ws.Dimension.Rows                                          | Should     be ($rowcount + 1)
        }
    }

    Context "#Example 2      # Exports a list of numbers and applies number format " {

        $path = "$env:TEMP\Test.xlsx"
        Remove-item -Path $path -ErrorAction SilentlyContinue
        #testing -ReturnRange switch
        $returnedRange = Write-Output -1 668 34 777 860 -0.5 119 -0.1 234 788,"=A9+A10" | Export-Excel -NumberFormat '[Blue]$#,##0.00;[Red]-$#,##0.00' -Path $path -ReturnRange
        it "Created a new file and returned the expected range                                     " {
            Test-Path -Path $path -ErrorAction SilentlyContinue         | Should     be $true
            $returnedRange                                              | Should     be "A1:A11"
        }

        $Excel = Open-ExcelPackage -Path $path
        it "Created 1 worksheet                                                                    " {
            $Excel.Workbook.Worksheets.count                            | Should     be 1
        }

        $ws = $Excel.Workbook.Worksheets[1]
        it "Created the worksheet with the expected name, number of rows and number of columns     " {
            $ws.Name                                                    | Should     be "sheet1"
            $ws.Dimension.Columns                                       | Should     be  1
            $ws.Dimension.Rows                                          | Should     be  11
        }

        it "Set the default style for the sheet as expected                                        " {
            $ws.cells.Style.Numberformat.Format                         | Should     be  '[Blue]$#,##0.00;[Red]-$#,##0.00'
        }

        it "Set the default style and value for Cell A1  as expected                               " {
            $ws.cells[1, 1].Style.Numberformat.Format                   | Should     be  '[Blue]$#,##0.00;[Red]-$#,##0.00'
            $ws.cells[1, 1].Value                                       | Should     be  -1
        }
    }

    Context "#Examples 3 & 4 # Setting cells for different data types Also added test for URI type" {

        if ((Get-Culture).NumberFormat.CurrencySymbol -eq "£") {$OtherCurrencySymbol = "$"}
        else {$OtherCurrencySymbol = "£"}
        $path = "$env:TEMP\Test.xlsx"
        Remove-item -Path $path -ErrorAction SilentlyContinue
        [PSCustOmobject][Ordered]@{
            Date             = Get-Date
            Formula1         = '=SUM(F2:G2)'
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
            strLocalCurrency = ('{0}123.45' -f (Get-Culture).NumberFormat.CurrencySymbol )
            strOtherCurrency = ('{0}123.45' -f $OtherCurrencySymbol )
            StrE164Phone     = '+32 (444) 444 4444'
            StrAltPhone1     = '+32 4 4444 444'
            StrAltPhone2     = '+3244444444'
            StrLeadSpace    = '  123'
            StrTrailSpace   = '123   '
            Link1            = [uri]"https://github.com/dfinke/ImportExcel"
            Link2            = "https://github.com/dfinke/ImportExcel"
            Link3            = "xl://internal/sheet1!A1"
            Link4            = "xl://internal/sheet1!C5"
        } | Export-Excel  -NoNumberConversion IPAddress, StrLeadZero, StrAltPhone2  -Path $path
        it "Created a new file                                                                     " {
            Test-Path -Path $path -ErrorAction SilentlyContinue         | Should     be $true
        }
        $Excel = Open-ExcelPackage -Path $path
        it "Created 1 worksheet                                                                    " {
            $Excel.Workbook.Worksheets.count                            | Should     be 1
        }
        $ws = $Excel.Workbook.Worksheets[1]
        it "Created the worksheet with the expected name, number of rows and number of columns     " {
            $ws.Name                                                    | Should     be "sheet1"
            $ws.Dimension.Columns                                       | Should     be  24
            $ws.Dimension.Rows                                          | Should     be  2
        }
        it "Set a date     in Cell A2                                                              " {
            $ws.Cells[2, 1].Value.Gettype().name                        | Should     be  'DateTime'
        }
        it "Set a formula  in Cell B2                                                              " {
            $ws.Cells[2, 2].Formula                                     | Should     be  '=SUM(F2:G2)'
        }
        it "Set strings    in Cells E2, F2 and R2  (no number conversion)                          " {
            $ws.Cells[2,  5].Value.GetType().name                        | Should     be  'String'
            $ws.Cells[2,  6].Value.GetType().name                        | Should     be  'String'
            $ws.Cells[2, 18].Value.GetType().name                        | Should     be  'String'
        }
        it "Set numbers    in Cells K2,L2,M2   (diferent Negative integer formats)                 " {
            ($ws.Cells[2, 11].Value -is [valuetype] )                   | Should     be  $true
            ($ws.Cells[2, 12].Value -is [valuetype] )                   | Should     be  $true
            ($ws.Cells[2, 13].Value -is [valuetype] )                   | Should     be  $true
             $ws.Cells[2, 11].Value                                     | Should     beLessThan 0
             $ws.Cells[2, 12].Value                                     | Should     beLessThan 0
             $ws.Cells[2, 13].Value                                     | Should     beLessThan 0
        }
        it "Set external hyperlinks in Cells U2 and V2                                             " {
            $ws.Cells[2, 21].Hyperlink                                 | Should     be  "https://github.com/dfinke/ImportExcel"
            $ws.Cells[2, 22].Hyperlink                                 | Should     be  "https://github.com/dfinke/ImportExcel"
        }
        it "Set internal hyperlinks in Cells W2 and X2                                             " {
            $ws.Cells[2, 23].Hyperlink.Scheme                          | Should     be  "xl"
            $ws.Cells[2, 23].Hyperlink.ReferenceAddress                | Should     be  "sheet1!A1"
            $ws.Cells[2, 23].Hyperlink.Display                         | Should     be  "sheet1"
            $ws.Cells[2, 24].Hyperlink.Scheme                          | Should     be  "xl"
            $ws.Cells[2, 24].Hyperlink.ReferenceAddress                | Should     be  "sheet1!c5"
            $ws.Cells[2, 24].Hyperlink.Display                         | Should     be  "sheet1!c5"
        }
        it "Processed thousands according to local settings   (Cells H2 and I2)                    " {
            if ((Get-Culture).NumberFormat.NumberGroupSeparator -EQ ",") {
                ($ws.Cells[2, 8].Value -is [valuetype] )               | Should     be  $true
                 $ws.Cells[2, 9].Value.GetType().name                  | Should     be  'String'
            }
            elseif ((Get-Culture).NumberFormat.NumberGroupSeparator -EQ ".") {
                ($ws.Cells[2, 9].Value -is [valuetype] )               | Should     be  $true
                 $ws.Cells[2, 8].Value.GetType().name                  | Should     be  'String'
            }
        }
        it "Processed local currency as a number and other currency as a string (N2 & O2)          " {
            ($ws.Cells[2, 14].Value -is [valuetype] )                   | Should     be  $true
             $ws.Cells[2, 15].Value.GetType().name                      | Should     be  'String'
        }
        it "Processed numbers with spaces between digits as strings (P2 & Q2)                      " {
             $ws.Cells[2, 16].Value.GetType().name                      | Should     be  'String'
             $ws.Cells[2, 17].Value.GetType().name                      | Should     be  'String'
        }
        it "Processed numbers leading or trailing speaces as Numbers (S2 & T2)                     " {
            ($ws.Cells[2, 19].Value -is [valuetype] )                   | Should     be  $true
            ($ws.Cells[2, 20].Value -is [valuetype] )                   | Should     be  $true
        }
    }

    Context "#               # Setting cells for different data types with -noHeader" {

        $path = "$env:TEMP\Test.xlsx"
        Remove-item -Path $path -ErrorAction SilentlyContinue
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
            Test-Path -Path $path -ErrorAction SilentlyContinue         | Should     be $true
        }

        $Excel = Open-ExcelPackage -Path $path
        it "Created 1 worksheet                                                                    " {
            $Excel.Workbook.Worksheets.count                            | Should     be 1
        }

        $ws = $Excel.Workbook.Worksheets[1]
        it "Created the worksheet with the expected name, number of rows and number of columns     " {
            $ws.Name                                                    | Should     be "sheet1"
            $ws.Dimension.Columns                                       | Should     be  14
            $ws.Dimension.Rows                                          | Should     be  1
        }

        it "Set a date      in Cell A1                                                             " {
            $ws.Cells[1, 1].Value.Gettype().name                         | Should     be  'DateTime'
        }

        it "Set a formula   in Cell B1                                                             " {
            $ws.Cells[1, 2].Formula                                      | Should     be  '=SUM(F1:G1)'
        }

        it "Set strings     in Cells E1 and F1                                                     " {
            $ws.Cells[1, 5].Value.GetType().name                        | Should     be  'String'
            $ws.Cells[1, 6].Value.GetType().name                        | Should     be  'String'
        }

        it "Set a number    in Cell I1                                                             " {
            ($ws.Cells[1, 9].Value -is [valuetype] )                     | Should     be  $true
        }

        it "Set a hyperlink in Cell N1                                                             " {
            $ws.Cells[1, 14].Hyperlink                                   | Should     be  "https://github.com/dfinke/ImportExcel"
        }
    }

    Context "#Example 5      # Adding a single conditional format " {
        $ct = New-ConditionalText -ConditionalType GreaterThan 525 -ConditionalTextColor DarkRed -BackgroundColor LightPink
        it "Created a Conditional format description                                               " {
            $ct.BackgroundColor -is [System.Drawing.Color]         | Should     be $true
            $ct.ConditionalTextColor -is [System.Drawing.Color]         | Should     be $true
            $ct.ConditionalType -in [enum]::GetNames( [OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType] ) |
                Should     be $true
        }

        $path = "$env:TEMP\Test.xlsx"
        Remove-item -Path $path -ErrorAction SilentlyContinue
        Write-Output 489 668 299 777 860 151 119 497 234 788 | Export-Excel -Path $path -ConditionalText $ct

        it "Created a new file                                                                     " {
            Test-Path -Path $path -ErrorAction SilentlyContinue         | Should     be $true
        }

        #ToDo need to test applying conitional formatting to a pre-existing worksheet
        $Excel = Open-ExcelPackage -Path $path
        $ws = $Excel.Workbook.Worksheets[1]

        it "Added one block of conditional formating for the data range                            " {
            $ws.ConditionalFormatting.Count                             | Should     be 1
            $ws.ConditionalFormatting[0].Address                        | Should     be ($ws.Dimension.Address)
        }

        $cf = $ws.ConditionalFormatting[0]
        it "Set the conditional formatting properties correctly                                    " {
            $cf.Formula                                                 | Should     be $ct.Text
            $cf.Type.ToString()                                         | Should     be $ct.ConditionalType
            #$cf.Style.Fill.BackgroundColor         | Should     be $ct.BackgroundColor
            # $cf.Style.Font.Color                   | Should be $ct.ConditionalTextColor  - have to compare r.g.b


        }
    }

    Context "#Example 6      # Adding multiple conditional formats using short form syntax. " {
        #this is a test of adding more than one conditional block and using the minimal syntax for new-ConditionalText =
        $path = "$env:TEMP\Test.xlsx"
        Remove-item -Path $path -ErrorAction SilentlyContinue

        #Testing -Passthrough
        $Excel = Get-Service | Select-Object Name, Status, DisplayName, ServiceName |
            Export-Excel $path -PassThru  -ConditionalText $(
            New-ConditionalText Stop DarkRed LightPink
            New-ConditionalText Running Blue Cyan
        )
        $ws = $Excel.Workbook.Worksheets[1]
        it "Added two blocks of conditional formating for the data range                           " {
            $ws.ConditionalFormatting.Count                             | Should     be 2
            $ws.ConditionalFormatting[0].Address                        | Should     be ($ws.Dimension.Address)
            $ws.ConditionalFormatting[1].Address                        | Should     be ($ws.Dimension.Address)
        }
        it "Set the conditional formatting properties correctly                                    " {
            $ws.ConditionalFormatting[0].Text                           | Should     be "Stop"
            $ws.ConditionalFormatting[1].Text                           | Should     be "Running"
            $ws.ConditionalFormatting[0].Type                           | Should     be "ContainsText"
            $ws.ConditionalFormatting[1].Type                           | Should     be "ContainsText"
            #Add RGB Comparison
        }
        Close-ExcelPackage -ExcelPackage $Excel
    }

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
        $newarray = $Array | Update-FirstObjectProperties
        it "Outputs as many objects as it input                                                    " {
            $newarray.Count                                             | Should     be $Array.Count
        }
        it "Added properties to item 0                                                             " {
            $newarray[0].psobject.Properties.name.Count                 | Should     be 4
            $newarray[0].Member1                                        | Should     be 'First'
            $newarray[0].Member2                                        | Should     be 'Second'
            $newarray[0].Member3                                        | Should     beNullOrEmpty
            $newarray[0].Member4                                        | Should     beNullOrEmpty
        }
    }

    Context "#Examples 8 & 9 # Adding Pivot tables and charts from parameters" {
        $path = "$env:TEMP\Test.xlsx"
        #This time we are not deleting the XLSX file so this Should create a new, named, sheet.
        $Excel = Get-Process |  Select-Object -first 50 -Property Name, cpu, pm, handles, company |  Export-Excel  $path -WorkSheetname Processes -PassThru
        #Testing -passthru and adding the Pivot as a second step. Want to save and re-open it ...
        Export-Excel -ExcelPackage $Excel -WorkSheetname Processes -IncludePivotTable -PivotRows Company -PivotData PM -NoTotalsInPivot

        $Excel = Open-ExcelPackage  $path
        $PTws = $Excel.Workbook.Worksheets["ProcessesPivotTable"]
        $wCount = $Excel.Workbook.Worksheets.Count
        it "Added the named sheet and pivot table to the workbook                                  " {
            $PTws                                                       | Should not beNullOrEmpty
            $PTws.PivotTables.Count                                     | Should     be 1
            $Excel.Workbook.Worksheets["Processes"]                     | Should not beNullOrEmpty
            $Excel.Workbook.Worksheets.Count                            | Should     beGreaterThan 2
            # $excel.Workbook.Worksheets["Processes"].Dimension.rows      | Should     be 51    #50 data + 1 header
        }
        $pt = $PTws.PivotTables[0]
        it "Built the expected Pivot table                                                         " {
            $pt.RowFields.Count                                         | Should     be 1
            $pt.RowFields[0].Name                                       | Should     be "Company"
            $pt.DataFields.Count                                        | Should     be 1
            $pt.DataFields[0].Function                                  | Should     be "Count"
            $pt.DataFields[0].Field.Name                                | Should     be "PM"
            $PTws.Drawings.Count                                        | Should     be 0
        }
        #using the already open sheet add the pivot chart
        $warnvar = $null
        Export-Excel -ExcelPackage $Excel -WorkSheetname Processes -IncludePivotTable -PivotRows Company -PivotData PM -IncludePivotChart -ChartType PieExploded3D -ShowCategory -NoLegend -WarningAction SilentlyContinue -WarningVariable warnvar
        $Excel = Open-ExcelPackage   $path
        it "Added a chart to the pivot table without rebuilding                                    " {
            $ws = $Excel.Workbook.Worksheets["ProcessesPivotTable"]
            $Excel.Workbook.Worksheets.Count                            | Should     be $wCount
            $ws.Drawings.count                                          | Should     be 1
            $ws.Drawings[0].ChartType.ToString()                        | Should     be "PieExploded3D"
        }
        it "Generated a message on re-processing the Pivot table                                   " {
            $warnVar                                                    | Should not beNullOrEmpty
        }
        $warnVar = $null
        Get-Process |  Select-Object -Last 50 -Property Name, cpu, pm, handles, company |   Export-Excel  $path -WorkSheetname Processes -Append -IncludePivotTable -PivotRows Company -PivotData PM -IncludePivotChart -ChartType PieExploded3D -WarningAction SilentlyContinue -WarningVariable warnvar
        $Excel = Open-ExcelPackage   $path
        $pt = $Excel.Workbook.Worksheets["ProcessesPivotTable"].PivotTables[0]
        it "Appended to the Worksheet and Extended the Pivot table                                 " {
            $Excel.Workbook.Worksheets.Count                            | Should     be $wCount
            # $excel.Workbook.Worksheets["Processes"].Dimension.rows      | Should     be 101     #appended 50 rows to the previous total
            # $pt.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref |
            #     Should     be "A1:E101"
        }
        it "Generated a message on extending the Pivot table                                       " {
            $warnVar                                                    | Should not beNullOrEmpty
        }
    }

    Context "                # Add-Worksheet inserted sheets, moved them correctly, and copied a sheet" {
        $path = "$env:TEMP\Test.xlsx"

        $Excel = Open-ExcelPackage  $path
        #At this point Sheets Should be in the order Sheet1, Processes, ProcessesPivotTable
        $null = Add-WorkSheet -ExcelPackage $Excel -WorkSheetname "Processes" -MoveToEnd   # order now  Sheet1, ProcessesPivotTable, Processes
        $null = Add-WorkSheet -ExcelPackage $Excel -WorkSheetname "NewSheet"  -MoveAfter "*" -CopySource ($excel.Workbook.Worksheets["Sheet1"]) # Now its NewSheet, Sheet1, ProcessesPivotTable, Processes
        $null = Add-WorkSheet -ExcelPackage $Excel -WorkSheetname "Sheet1"    -MoveAfter "Processes"  # Now its NewSheet, ProcessesPivotTable, Processes, Sheet1
        $null = Add-WorkSheet -ExcelPackage $Excel -WorkSheetname "Another"   -MoveToStart    # Now its Another, NewSheet, ProcessesPivotTable, Processes, Sheet1
        $null = Add-WorkSheet -ExcelPackage $Excel -WorkSheetname "OneLast"   -MoveBefore "ProcessesPivotTable"   # Now its Another, NewSheet, Onelast, ProcessesPivotTable, Processes, Sheet1
        Close-ExcelPackage $Excel

        $Excel = Open-ExcelPackage  $path

        it "Got the Sheets in the right order                                                      " {
            $excel.Workbook.Worksheets[1].Name  | Should be "Another"
            $excel.Workbook.Worksheets[2].Name  | Should be "NewSheet"
            $excel.Workbook.Worksheets[3].Name  | Should be "Onelast"
            $excel.Workbook.Worksheets[4].Name  | Should be "ProcessesPivotTable"
            $excel.Workbook.Worksheets[5].Name  | Should be "Processes"
            $excel.Workbook.Worksheets[6].Name  | Should be "Sheet1"
        }

        it "Cloned 'Sheet1' to 'NewSheet'                                                          " {
            $newWs = $excel.Workbook.Worksheets["NewSheet"]
            $newWs.Dimension.Address                          | Should     be ($excel.Workbook.Worksheets["Sheet1"].Dimension.Address)
            $newWs.ConditionalFormatting.Count                | Should     be ($excel.Workbook.Worksheets["Sheet1"].ConditionalFormatting.Count)
            $newWs.ConditionalFormatting[0].Address.Address   | Should     be ($excel.Workbook.Worksheets["Sheet1"].ConditionalFormatting[0].Address.Address)
            $newWs.ConditionalFormatting[0].Formula           | Should     be ($excel.Workbook.Worksheets["Sheet1"].ConditionalFormatting[0].Formula)
        }

    }

    Context "                # Create and append with Start row and Start Column, inc ranges and Pivot table. " {
        $path = "$env:TEMP\Test.xlsx"
        #Catch warning
        $warnVar = $null
        #Test Append with no existing sheet. Test adding a named pivot table from a command line parameter
        Get-Process | Select-Object -first 10 -Property Name, cpu, pm, handles, company  | Export-Excel -StartRow 3 -StartColumn 3 -AutoFilter -AutoNameRange -BoldTopRow -IncludePivotTable  -PivotRows Company -PivotData PM -PivotTableName 'PTOffset' -Path $path -WorkSheetname withOffset -Append -PivotFilter Name -NoTotalsInPivot
        Get-Process | Select-Object -last  10 -Property Name, cpu, pm, handles, company  | Export-Excel -StartRow 3 -StartColumn 3 -AutoFilter -AutoNameRange -BoldTopRow -IncludePivotTable  -PivotRows Company -PivotData PM -PivotTableName 'PTOffset' -Path $path -WorkSheetname withOffset -Append -WarningAction SilentlyContinue -WarningVariable warnvar
        $Excel = Open-ExcelPackage   $path
        $dataWs = $Excel.Workbook.Worksheets["withOffset"]
        $pt = $Excel.Workbook.Worksheets["PTOffset"].PivotTables[0]
        it "Created and appended to a sheet offset from the top left corner                        " {
            $dataWs.Cells[1, 1].Value                                   | Should     beNullOrEmpty
            $dataWs.Cells[2, 2].Value                                   | Should     beNullOrEmpty
            $dataWs.Cells[3, 3].Value                                   | Should not beNullOrEmpty
            $dataWs.Cells[3, 3].Style.Font.Bold                         | Should     be $true
            $dataWs.Dimension.End.Row                                   | Should     be 23
            $dataWs.names[0].end.row                                    | Should     be 23
            $dataWs.names[0].name                                       | Should     be 'Name'
            $dataWs.names.Count                                         | Should     be 6
            $dataWs.cells[$dataws.Dimension].AutoFilter                 | Should     be true
            $pt.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref |
                Should     be "C3:G23"
        }
        it "Generated a message on extending the Pivot table                                       " {
            $warnVar                                                    | Should not beNullOrEmpty
        }
    }

    Context "#Example 11     # Create and append with title, inc ranges and Pivot table" {
        $path = "$env:TEMP\Test.xlsx"
        $ptDef = [ordered]@{}
        $ptDef += New-PivotTableDefinition -PivotTableName "PT1" -SourceWorkSheet 'Sheet1' -PivotRows "Status"  -PivotData @{'Status'  = 'Count'} -PivotFilter "StartType" -IncludePivotChart -ChartType BarClustered3D  -ChartTitle "Services by status" -ChartHeight 512 -ChartWidth 768 -ChartRow 10 -ChartColumn 0 -NoLegend
        $ptDef += New-PivotTableDefinition -PivotTableName "PT2" -SourceWorkSheet 'Sheet2' -PivotRows "Company" -PivotData @{'Company' = 'Count'}                          -IncludePivotChart -ChartType PieExploded3D -ShowPercent -WarningAction SilentlyContinue

        it "Built a pivot definition using New-PivotTableDefinition                                " {
            $ptDef.PT1.SourceWorkSheet                                  | Should be 'Sheet1'
            $ptDef.PT1.PivotRows                                        | Should be 'Status'
            $ptDef.PT1.PivotData.Status                                 | Should be 'Count'
            $ptDef.PT1.PivotFilter                                      | Should be 'StartType'
            $ptDef.PT1.IncludePivotChart                                | Should be  $true
            $ptDef.PT1.ChartType.tostring()                             | Should be 'BarClustered3D'
        }
        Remove-Item -Path $path
        #Catch warning
        $warnvar = $null
        Get-Service | Select-Object    -Property Status, Name, DisplayName, StartType | Export-Excel -Path $path  -AutoSize                         -TableName "All Services"  -TableStyle Medium1 -WarningAction SilentlyContinue -WarningVariable warnvar
        Get-Process | Select-Object    -Property Name, Company, Handles, CPU, VM      | Export-Excel -Path $path  -AutoSize -WorkSheetname 'sheet2' -TableName "Processes"     -TableStyle Light1 -Title "Processes" -TitleFillPattern Solid -TitleBackgroundColor AliceBlue -TitleBold -TitleSize 22 -PivotTableDefinition $ptDef
        $Excel = Open-ExcelPackage   $path
        $ws1 = $Excel.Workbook.Worksheets["Sheet1"]
        $ws2 = $Excel.Workbook.Worksheets["Sheet2"]


        it "Set Column widths (with autosize)                                                      " {
            $ws1.Column(2).Width                                        | Should not be $ws1.DefaultColWidth
            $ws2.Column(1).width                                        | Should not be $ws2.DefaultColWidth
        }

        it "Added tables to both sheets (handling illegal chars) and a title in sheet 2            " {
            $warnvar.count                                              | Should     be 1
            $ws1.tables.Count                                           | Should     be 1
            $ws2.tables.Count                                           | Should     be 1
            $ws1.Tables[0].Address.Start.Row                            | Should     be 1
            $ws2.Tables[0].Address.Start.Row                            | Should     be 2 #Title in row 1
            $ws1.Tables[0].Address.End.Address                          | Should     be $ws1.Dimension.End.Address
            $ws2.Tables[0].Address.End.Address                          | Should     be $ws2.Dimension.End.Address
            $ws2.Tables[0].Name                                         | Should     be "Processes"
            $ws2.Tables[0].StyleName                                    | Should     be "TableStyleLight1"
            $ws2.Cells["A1"].Value                                      | Should     be "Processes"
            $ws2.Cells["A1"].Style.Font.Bold                            | Should     be $true
            $ws2.Cells["A1"].Style.Font.Size                            | Should     be 22
            $ws2.Cells["A1"].Style.Fill.PatternType.tostring()          | Should     be "solid"
            $ws2.Cells["A1"].Style.Fill.BackgroundColor.Rgb             | Should     be "fff0f8ff"
        }

        $ptsheet1 = $Excel.Workbook.Worksheets["Pt1"]
        $ptsheet2 = $Excel.Workbook.Worksheets["Pt2"]
        $PT1 = $ptsheet1.PivotTables[0]
        $PT2 = $ptsheet2.PivotTables[0]
        $PC1 = $ptsheet1.Drawings[0]
        $PC2 = $ptsheet2.Drawings[0]
        it "Created the correct pivot tables and charts from the definitions.                      " {

            $PT1.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref |
                Should     be ("A1:" + $ws1.Dimension.End.Address)
            $PT2.CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref |
                Should     be ("A2:" + $ws2.Dimension.End.Address) #Title in row 1

            $pt1.PageFields[0].Name                                     | Should     be 'StartType'
            $pt1.RowFields[0].Name                                      | Should     be 'Status'
            $pt1.DataFields[0].Field.name                               | Should     be 'Status'
            $pt1.DataFields[0].Function                                 | Should     be 'Count'
            $pc1.ChartType                                              | Should     be 'BarClustered3D'
            $pc1.From.Column                                            | Should     be 0                    #chart 1 at 0,10 chart 2 at 4,0 (default)
            $pc2.From.Column                                            | Should     be 4
            $pc1.From.Row                                               | Should     be 10
            $pc2.From.Row                                               | Should     be 0
            $pc1.Legend.Font                                            | Should     beNullOrEmpty           #Best check for legend removed.
            $pc2.Legend.Font                                            | Should not beNullOrEmpty
            $pc1.Title.Text                                             | Should     be 'Services by status'
            $pc2.DataLabel.ShowPercent                                  | Should     be $true
        }
    }

    Context "#Example 13     # Formatting and another way to do a pivot.  " {
        $path = "$env:TEMP\Test.xlsx"
        Remove-Item $path
        $excel = Get-Process | Select-Object -Property Name, Company, Handles, CPU, PM, NPM, WS | Export-Excel -Path $path -ClearSheet -WorkSheetname "Processes" -FreezeTopRowFirstColumn -PassThru
        $sheet = $excel.Workbook.Worksheets["Processes"]
        $sheet.Column(1) | Set-Format -Bold -AutoFit
        $sheet.Column(2) | Set-Format -Width 29 -WrapText
        $sheet.Column(3) | Set-Format -HorizontalAlignment Right -NFormat "#,###"
        Set-Format -Address $sheet.Cells["E1:H1048576"]  -HorizontalAlignment Right -NFormat "#,###"
        Set-Format -Address $sheet.Column(4)  -HorizontalAlignment Right -NFormat "#,##0.0" -Bold
        Set-Format -Address $sheet.Row(1) -Bold -HorizontalAlignment Center
        Add-ConditionalFormatting -WorkSheet $sheet -Range "D2:D1048576" -DataBarColor Red
        $rule = Add-ConditionalFormatting -passthru -Address $sheet.cells["C:C"] -RuleType TopPercent -ConditionValue 20 -Bold -StrikeThru
        Add-ConditionalFormatting -WorkSheet $sheet -Range "G2:G1048576" -RuleType GreaterThan -ConditionValue "104857600" -ForeGroundColor Red -Bold -Italic -Underline -BackgroundColor Beige -BackgroundPattern LightUp -PatternColor Gray
        foreach ($c in 5..9) {Set-Format $sheet.Column($c)  -AutoFit }
        Add-PivotTable -PivotTableName "PT_Procs" -ExcelPackage $excel -SourceWorkSheet 1 -PivotRows Company -PivotData  @{'Name' = 'Count'} -IncludePivotChart -ChartType ColumnClustered -NoLegend
        Close-ExcelPackage $excel

        $excel = Open-ExcelPackage $path
        $sheet = $excel.Workbook.Worksheets["Processes"]
        it "Returned the rule when calling Add-ConditionalFormatting -passthru                     " {
            $rule                                                       | should not beNullOrEmpty
            $rule.getType().fullname                                    | should     be "OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingTopPercent"
            $rule.Style.Font.Strike                                     | should be true
        }
        it "Applied the formating                                                                  " {
            $sheet                                                      | Should not beNullOrEmpty
            $sheet.Column(1).wdith                                      | Should not be  $sheet.DefaultColWidth
            $sheet.Column(7).wdith                                      | Should not be  $sheet.DefaultColWidth
            $sheet.Column(1).style.font.bold                            | Should     be  $true
            $sheet.Column(2).style.wraptext                             | Should     be  $true
            $sheet.Column(2).width                                      | Should     be  29
            $sheet.Column(3).style.horizontalalignment                  | Should     be  'right'
            $sheet.Column(4).style.horizontalalignment                  | Should     be  'right'
            $sheet.Cells["A1"].Style.HorizontalAlignment                | Should     be  'Center'
            $sheet.Cells['E2'].Style.HorizontalAlignment                | Should     be  'right'
            $sheet.Cells['A1'].Style.Font.Bold                          | Should     be  $true
            $sheet.Cells['D2'].Style.Font.Bold                          | Should     be  $true
            $sheet.Cells['E2'].style.numberformat.format                | Should     be  '#,###'
            $sheet.Column(3).style.numberformat.format                  | Should     be  '#,###'
            $sheet.Column(4).style.numberformat.format                  | Should     be  '#,##0.0'
            $sheet.ConditionalFormatting.Count                          | Should     be  3
            $sheet.ConditionalFormatting[0].type                        | Should     be  'Databar'
            $sheet.ConditionalFormatting[0].Color.name                  | Should     be  'ffff0000'
            $sheet.ConditionalFormatting[0].Address.Address             | Should     be  'D2:D1048576'
            $sheet.ConditionalFormatting[1].Style.Font.Strike           | Should     be  $true
            $sheet.ConditionalFormatting[1].type                        | Should     be  "TopPercent"
            $sheet.ConditionalFormatting[2].type                        | Should     be  'GreaterThan'
            $sheet.ConditionalFormatting[2].Formula                     | Should     be  '104857600'
            $sheet.ConditionalFormatting[2].Style.Font.Color.Color.Name | Should     be  'ffff0000'
        }
        it "Froze the panes                                                                        " {
            $sheet.view.Panes.Count                                     | Should     be 3
        }
        $ptsheet1 = $Excel.Workbook.Worksheets["Pt_procs"]

        it "Created the pivot table                                                                " {
            $ptsheet1                                                   | Should not beNullOrEmpty
            $ptsheet1.PivotTables[0].DataFields[0].Field.Name           | Should     be "Name"
            $ptsheet1.PivotTables[0].DataFields[0].Function             | Should     be "Count"
            $ptsheet1.PivotTables[0].RowFields[0].Name                  | Should     be "Company"
            $ptsheet1.PivotTables[0].CacheDefinition.CacheDefinitionXml.pivotCacheDefinition.cacheSource.worksheetSource.ref |
                Should     be $sheet.Dimension.address
        }
    }

    Context "                # Chart from MultiSeries.ps1 in the Examples\charts Directory" {
        $path = "$env:TEMP\Test.xlsx"
        Remove-Item -Path   $path -ErrorAction SilentlyContinue
        $data = invoke-sum (Get-Process) Company Handles, PM, VirtualMemorySize
        it "used Invoke-Sum to create a data set                                                   " {
            $data                                                       | Should not beNullOrEmpty
            $data.count                                                 | Should     beGreaterThan 1
            $data[1].Name                                               | Should  not beNullOrEmpty
            $data[1].Handles                                            | Should  not beNullOrEmpty
            $data[1].PM                                                 | Should  not beNullOrEmpty
            $data[1].VirtualMemorySize                                  | Should  not beNullOrEmpty
        }
        $c = New-ExcelChartDefinition -Title Stats -ChartType LineMarkersStacked   -XRange "Processes[Name]" -YRange "Processes[PM]", "Processes[VirtualMemorySize]" -SeriesHeader 'PM', 'VMSize'

        it "Created the Excel chart definition                                                     " {
            $c                                                          | Should not beNullOrEmpty
            $c.ChartType.gettype().name                                 | Should     be "eChartType"
            $c.ChartType.tostring()                                     | Should     be "LineMarkersStacked"
            $c.yrange -is [array]                                       | Should     be $true
            $c.yrange.count                                             | Should     be 2
            $c.yrange[0]                                                | Should     be "Processes[PM]"
            $c.yrange[1]                                                | Should     be "Processes[VirtualMemorySize]"
            $c.xrange                                                   | Should     be "Processes[Name]"
            $c.Title                                                    | Should     be "Stats"
            $c.Nolegend                                                 | Should not be $true
            $c.ShowCategory                                             | Should not be $true
            $c.ShowPercent                                              | Should not be $true
        }
        $data | Export-Excel $path -AutoSize -TableName Processes -ExcelChartDefinition $c
        $excel = Open-ExcelPackage -Path $path
        $drawings = $excel.Workbook.Worksheets[1].drawings
        it "Used the Excel chart definition with Export-Excel                                      " {
            $drawings.count                                             | Should     be 1
            $drawings[0].ChartType                                      | Should     be "LineMarkersStacked"
            $drawings[0].Series.count                                   | Should     be 2
            $drawings[0].Series[0].Series                               | Should     be "'Sheet1'!Processes[PM]"
            $drawings[0].Series[0].XSeries                              | Should     be "'Sheet1'!Processes[Name]"
            $drawings[0].Series[1].Series                               | Should     be "'Sheet1'!Processes[VirtualMemorySize]"
            $drawings[0].Series[1].XSeries                              | Should     be "'Sheet1'!Processes[Name]"
            $drawings[0].Title.text                                     | Should     be "Stats"
        }
        Close-ExcelPackage $excel
    }

    Context "                # variation of plot.ps1 from Examples Directory using Add chart outside ExportExcel" {
        $path = "$env:TEMP\Test.xlsx"
        $excel = 0..360 | ForEach-Object {[pscustomobject][ordered]@{x = $_; Sinx = "=Sin(Radians(x)) "}} | Export-Excel -AutoNameRange -Path $path -WorkSheetname SinX -ClearSheet -FreezeFirstColumn -PassThru
        Add-ExcelChart -Worksheet $excel.Workbook.Worksheets["Sinx"] -ChartType line -XRange "X" -YRange "Sinx" -SeriesHeader "Sin(x)" -Title "Graph of Sine X" -TitleBold -TitleSize 14 `
                      -Column 2 -ColumnOffSetPixels 35 -Width 800 -XAxisTitleText "Degrees" -XAxisTitleBold -XAxisTitleSize 12 -XMajorUnit 30 -XMinorUnit 10 -XMinValue 0 -XMaxValue 361  -XAxisNumberformat "000" -XAxisPosition Bottom `
                      -YMinValue -1.25 -YMaxValue 1.25 -YMajorUnit 0.25 -YAxisNumberformat "0.00" -YAxisTitleText "Sine" -YAxisTitleBold -YAxisTitleSize 12 -YAxisPosition Left -LegendPostion Bottom -LegendSize 8 -legendBold
        $d = $excel.Workbook.Worksheets["Sinx"].Drawings[0]
        It "Controled the axes and title and legend of the chart                                   " {
            $d.XAxis.MaxValue                                           | Should     be 361
            $d.XAxis.MajorUnit                                          | Should     be 30
            $d.XAxis.MinorUnit                                          | Should     be 10
            $d.XAxis.Title.Text                                         | Should     be "degrees"
            $d.XAxis.Title.Font.bold                                    | Should     be $true
            $d.XAxis.Title.Font.Size                                    | Should     be 12
            $d.XAxis.MajorUnit                                          | Should     be 30
            $d.XAxis.MinorUnit                                          | Should     be 10
            $d.XAxis.MinValue                                           | Should     be 0
            $d.XAxis.MaxValue                                           | Should     be 361
            $d.YAxis.Format                                             | Should     be "0.00"
            $d.Title.Text                                               | Should     be "Graph of Sine X"
            $d.Title.Font.Bold                                          | Should     be $true
            $d.Title.Font.Size                                          | Should     be 14
            $d.yAxis.MajorUnit                                          | Should     be 0.25
            $d.yAxis.MaxValue                                           | Should     be 1.25
            $d.yaxis.MinValue                                           | Should     be -1.25
            $d.Legend.Position.ToString()                               | Should     be "Bottom"
            $d.Legend.Font.Bold                                         | Should     be $true
            $d.Legend.Font.Size                                         | Should     be 8
            $d.ChartType.tostring()                                     | Should     be "line"
            $d.From.Column                                              | Should     be 2
        }
        Close-ExcelPackage -ExcelPackage $excel -nosave
    }

    Context "                # Quick Pie chart and three icon conditional formating" {
        $path = "$Env:TEMP\Pie.xlsx"
        Remove-Item -Path $path -ErrorAction SilentlyContinue

        $range = Get-Process| Group-Object -Property company | Where-Object -Property name |
             Select-Object -Property Name, @{n="TotalPm";e={($_.group | Measure-Object -sum -Property pm).sum }} |
                 Export-Excel -NoHeader -AutoNameRange -path $path -ReturnRange  -PieChart -ShowPercent
        $Cf = New-ConditionalFormattingIconSet -Range ($range -replace "^.*:","B2:") -ConditionalFormat ThreeIconSet -Reverse -IconType Flags
        $ct = New-ConditionalText -Text "Microsoft" -ConditionalTextColor red -BackgroundColor AliceBlue -ConditionalType ContainsText
        it "Created the Conditional formatting rules                                               " {
            $cf.Formatter                                               | should     be "ThreeIconSet"
            $cf.IconType                                                | should     be "Flags"
            $cf.Range                                                   | Should     be ($range -replace "^.*:","B2:")
            $cf.Reverse                                                 | Should     be $true
            $ct.BackgroundColor.Name                                    | Should     be "AliceBlue"
            $ct.ConditionalTextColor.Name                               | Should     be "Red"
            $ct.ConditionalType                                         | Should     be "ContainsText"
            $ct.Text                                                    | Should     be "Microsoft"
        }

        Export-Excel -Path $path -ConditionalFormat $cf -ConditionalText $ct
        $excel = Open-ExcelPackage -Path $path
        $rows  = $range -replace "^.*?(\d+)$", '$1'
        $chart = $excel.Workbook.Worksheets["sheet1"].Drawings[0]
        $cFmt  = $excel.Workbook.Worksheets["sheet1"].ConditionalFormatting
        it "Created the chart with the right series                                                " {
            $chart.ChartType                                            | should     be "PieExploded3D"
            $chart.series.series                                        | should     be "'Sheet1'!B1:B$rows" #would be B2 and A2 if we had a header.
            $chart.series.Xseries                                       | should     be "'Sheet1'!A1:A$rows"
            $chart.DataLabel.ShowPercent                                | should     be $true
        }
        it "Created two Conditional formatting rules                                               " {
            $cFmt.Count                                                 | should     be $true
            $cFmt.Where({$_.type -eq "ContainsText"})                   | Should not beNullOrEmpty
            $cFmt.Where({$_.type -eq "ThreeIconSet"})                   | Should not beNullOrEmpty
        }
    }

    Context "                # Awkward multiple tables" {
        $path = "$Env:TEMP\test.xlsx"
        remove-item -Path $path -ErrorAction SilentlyContinue
        $r = Get-ChildItem -path C:\WINDOWS\system32 -File

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
            $ws.tables.count | should be 3
        }
        it "Created the FileSize table in the right place with the right size and style            " {
            $ws.Tables["FileSize"].Address.Address                      | should     be "G2:H16" #Insert at row 2, Column 7, 14 rows x 2 columns of data
            $ws.Tables["FileSize"].StyleName                            | should     be "TableStyleMedium2"
        }
        it "Created the ExtSize  table in the right place with the right size                      " {
            $ws.Tables["ExtSize"].Address.Address                      | should     be "A2:B14" #tile, then 12 rows x 2 columns of data
        }
        it "Created the ExtCount table in the right place with the right size                      " {
            $ws.Tables["ExtCount"].Address.Address                      | should     be "D2:E12" #title, then 10 rows x 2 columns of data
        }
    }

}
