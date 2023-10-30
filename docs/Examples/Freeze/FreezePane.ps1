# Freeze the columns/rows to left and above the cell

$data = ConvertFrom-Csv @"
Region,State,Units,Price,Name,NA,EU,JP,Other
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
West,Texas,927,923.71,Wii Sports,41.49,29.02,3.77,8.46
"@

$xlfilename = "test.xlsx"
Remove-Item $xlfilename -ErrorAction SilentlyContinue

<#
    Freezes the top two rows and the two leftmost column
#>

$data | Export-Excel $xlfilename -Show -Title 'Sales Data' -FreezePane 3, 3