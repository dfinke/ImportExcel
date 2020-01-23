function Expand-NumberFormat {
    [CmdletBinding()]
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
