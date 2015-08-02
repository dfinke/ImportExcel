cls

Add-Type -Path .\EPPlus.dll

$fileName = "c:\temp\TestCC.xlsx"

$r=ps excel -ErrorAction Ignore 
if($r) {
    $r|kill
    sleep -Milliseconds 250
}

rm $fileName -Force -ErrorAction Ignore

function get-rm { (Get-Random -Minimum 0 -Maximum 12) + 1}
function get-rd { (Get-Random -Minimum 0 -Maximum 27) + 1}

$pkg = New-Object OfficeOpenXml.ExcelPackage $fileName
$workbook = $pkg.Workbook

$ws = $pkg.Workbook.Worksheets.Add("Test Conditional Formatting")

foreach($col in 1..10) {
    $ws.Cells[1,$col].Value = "Sample $col"
    
    foreach($row in 2..10) {
        $ws.Cells[$row,$col].Value = [int] (Get-Random -Maximum 20 -Minimum 1)
    }
}

#$ws.Cells[1,1].Value = "State"
#$ws.Cells[1,2].Value = "Name"

#$ws.Cells[2,1].Value = "CA"
#$ws.Cells[2,2].Value = "California"

#$ws.Cells[3,1].Value = "NY"
#$ws.Cells[3,2].Value = "New York"

#$ws.Cells[4,1].Value = "IL"
#$ws.Cells[4,2].Value = "Illinois"

#$ws.Cells[5,1].Value = "NH"
#$ws.Cells[5,2].Value = "New Hampshire"


#$rule = $ws.ConditionalFormatting.AddNotContainsText("B2:B5")
#$rule.Style.Fill.BackgroundColor.Color = "red"
#$rule.Text = 'New'
#$rule.Formula = 'New'

#$rule = $ws.ConditionalFormatting.AddEqual("B:B")
#$rule.Style.Fill.BackgroundColor.Color = "red"
#$rule.Formula = '"New York"'

#$rule = $ws.ConditionalFormatting.AddEqual("B:B")
#$rule.Style.Fill.BackgroundColor.Color = "Cyan"
#$rule.Formula = '"Illinois"'


#foreach($col in 1..10) {
#    $ws.Cells[1,$col].Value = "Sample $col"
#    
#    foreach($row in 2..10) {
#        
#        if($row -eq 2 -or $row -eq 7) {
#            $ws.Cells[$row,$col].Value = get-date
#        } else {
#            $ws.Cells[$row,$col].Value = get-date ("{0}/{1}/2015" -f (get-rm), (get-rd))
#        }
#
#        $ws.Cells[$row,$col].Style.Numberformat.Format = "m/d/yyyy"
#    }
#}

#$rule.Style.Font.Bold = $true

#$rule = $ws.ConditionalFormatting.AddEqual("A2:J10")
#$rule.Style.Fill.BackgroundColor.Color = "red"
#$rule.Formula = "6"

#$rule = $ws.ConditionalFormatting.AddEqual("A2:J10")
#$rule.Style.Fill.BackgroundColor.Color = "cyan"
#$rule.Formula = "4"


#$r = $ws.ConditionalFormatting.AddThisWeek("A2:J10")
#$r.Style.Font.Bold   = $true
#$r.Style.Font.Italic = $true
#$r.Style.Font.Color = [OfficeOpenXml.Style.Dxf.ExcelDxfColor]"Red"

#$range = New-Object OfficeOpenXml.ExcelAddress "B2:B10"
#$range = New-Object OfficeOpenXml.ExcelAddress "B2:B10"
#$r = $ws.ConditionalFormatting.AddThreeColorScale("B2:B10")
#$r = $ws.ConditionalFormatting.AddThreeColorScale("5:5")
#$r.MiddleValue.Type  = "Percentile"
#$r.MiddleValue.Value = 30
#$r.StopIfTrue=$true

#$range = New-Object OfficeOpenXml.ExcelAddress "A2:A10"

#$IconSet = [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting3IconsSetType]::Flags
#$IconSet = [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting4IconsSetType]::RedToBlack

function Set-ConditionalFormatting {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelWorksheet]$targetWS,
        $Address,
        [ValidateSet("ThreeIconSet","FourIconSet","FiveIconSet")]
        $ConditionalFormat        
    )
    
    DynamicParam {       
        $IconType = New-Object System.Management.Automation.ParameterAttribute
        $IconType.Position = 2
        $IconType.Mandatory = $true

        $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
                
        $attributeCollection.Add($IconType)        

        switch ($ConditionalFormat) {
            "ThreeIconSet" {
                $IconTypeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('IconType', [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting3IconsSetType], $attributeCollection)                
            }

            "FourIconSet" {
                $IconTypeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('IconType', [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting4IconsSetType], $attributeCollection)                
            }

            "FiveIconSet" {
                $IconTypeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('IconType', [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting5IconsSetType], $attributeCollection)                
            }
        }

        $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        
        $paramDictionary.Add('IconType', $IconTypeParam)
        
        return $paramDictionary
    }
    
    Begin {
        $bp = @{}+$PSBoundParameters        
    }
    Process {
        $target = "Add$($ConditionalFormat)"
        $null = ($ws.ConditionalFormatting).$target($Address, $bp.IconType)
    }
}

$ws | Set-ConditionalFormatting -Address '3:3' -ConditionalFormat ThreeIconSet -IconType Symbols
#$rulex = $ws.ConditionalFormatting.AddFourIconSet("B:C", 'RedToBlack')
#$rulex.Reverse = $true


# $rulex = $ws.ConditionalFormatting.AddThreeIconSet("a2:a10", "Flags")
# $rulex = $ws.ConditionalFormatting.AddThreeIconSet("b2:b10", "Flags")
#$r = $ws.ConditionalFormatting.AddThreeIconSet("D:D", "Flags")
#$r = $ws.ConditionalFormatting.AddThreeIconSet("3:3", "Flags")
#$r = $ws.ConditionalFormatting.AddThreeIconSet("6:6", "Flags")

#$r = $ws.ConditionalFormatting.AddThreeIconSet("6:7", "Flags")
#$r = $ws.ConditionalFormatting.AddThreeIconSet("B:D", "Flags")

#$rule2 = $ws.ConditionalFormatting.AddDatabar("b2:b10", "blue")
#$rule3 = $ws.ConditionalFormatting.AddNotEqual("c2:c10")
#$rule3.Formula = "14"

#$ruleX = $ws.ConditionalFormatting.AddThreeColorScale("a2:a10")

#$cfRule1 = $ws.ConditionalFormatting.AddTwoColorScale("b2:b10")
#$cfRule1 = $ws.ConditionalFormatting.AddBelowAverage("B11:B20")

#$cfRule1.LowValue.Type  = [OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType]::Num
#$cfRule1.LowValue.Type  = 'Num'
#$cfRule1.LowValue.Value = 2
#$cfRule1.LowValue.Color = "Red" 

$pkg.Save()
$pkg.Dispose()

ii $fileName
