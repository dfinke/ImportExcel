
function Expand-NumberFormat {
    <#
      .SYNOPSIS
        Converts short names for number formats to the formatting strings used in Excel
      .DESCRIPTION
        Where you can type a number format you can write, for example, 'Short-Date'
        and the module will translate it into the format string used by Excel.
        Some formats, like Short-Date change how they are presented when Excel
        loads (so date will use the local ordering of year, month and Day). Other
        formats change how they appear when loaded with different cultures
        (depending on the country "," or "." or " " may be the thousand seperator
        although Excel always stores it as ",")
      .EXAMPLE
        Expand-NumberFormat percentage

        Returns "0.00%"
      .EXAMPLE
        Expand-NumberFormat Currency

        Returns the currency format specified in the local regional settings. This
        may not be the same as Excel uses.  The regional settings set the currency
        symbol and then whether it is before or after the number and separated with
        a space or not; for negative numbers the number may be wrapped in parentheses
        or a - sign might appear before or after the number and symbol.
        So this returns $#,##0.00;($#,##0.00) for English US, #,##0.00 €;€#,##0.00-
        for French. (Note some Eurozone countries write €1,23 and others 1,23€ )
        In French the decimal point will be rendered as a "," and the thousand
        separator as a space.
    #>
    [cmdletbinding()]
    [OutputType([String])]
    param  (
        #the format string to Expand
        $NumberFormat
    )
    switch ($NumberFormat) {
        "Currency"      {
            #https://msdn.microsoft.com/en-us/library/system.globalization.numberformatinfo.currencynegativepattern(v=vs.110).aspx
            $sign = [cultureinfo]::CurrentCulture.NumberFormat.CurrencySymbol
            switch ([cultureinfo]::CurrentCulture.NumberFormat.CurrencyPositivePattern) {
                0  {$pos = "$Sign#,##0.00"  ; break }
                1  {$pos = "#,##0.00$Sign"  ; break }
                2  {$pos = "$Sign #,##0.00" ; break }
                3  {$pos = "#,##0.00 $Sign" ; break }
            }
            switch ([cultureinfo]::CurrentCulture.NumberFormat.CurrencyPositivePattern) {
                0  {return "$pos;($Sign#,##0.00)"  }
                1  {return "$pos;-$Sign#,##0.00"   }
                2  {return "$pos;$Sign-#,##0.00"   }
                3  {return "$pos;$Sign#,##0.00-"   }
                4  {return "$pos;(#,##0.00$Sign)"  }
                5  {return "$pos;-#,##0.00$Sign"   }
                6  {return "$pos;#,##0.00-$Sign"   }
                7  {return "$pos;#,##0.00$Sign-"   }
                8  {return "$pos;-#,##0.00 $Sign"  }
                9  {return "$pos;-$Sign #,##0.00"  }
               10  {return "$pos;#,##0.00 $Sign-"  }
               11  {return "$pos;$Sign #,##0.00-"  }
               12  {return "$pos;$Sign -#,##0.00"  }
               13  {return "$pos;#,##0.00- $Sign"  }
               14  {return "$pos;($Sign #,##0.00)" }
               15  {return "$pos;(#,##0.00 $Sign)" }
            }
        }
        "Number"        {return  "0.00"       } # format id  2
        "Percentage"    {return  "0.00%"      } # format id 10
        "Scientific"    {return  "0.00E+00"   } # format id 11
        "Fraction"      {return  "# ?/?"      } # format id 12
        "Short Date"    {return  "mm-dd-yy"   } # format id 14 localized on load by Excel.
        "Short Time"    {return  "h:mm"       } # format id 20 localized on load by Excel.
        "Long Time"     {return  "h:mm:ss"    } # format id 21 localized on load by Excel.
        "Date-Time"     {return  "m/d/yy h:mm"} # format id 22 localized on load by Excel.
        "Text"          {return  "@"          } # format ID 49
        Default         {return  $NumberFormat}
    }
}
