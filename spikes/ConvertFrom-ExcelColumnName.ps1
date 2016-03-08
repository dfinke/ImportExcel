function ConvertFrom-ExcelColumnName {
    param($columnName)

    $sum=0
    $columnName.ToCharArray() |
        ForEach {
            $sum*=26
            $sum+=[char]$_.tostring().toupper()-[char]'A'+1
        } 
    $sum
}

ConvertFrom-ExcelColumnName p
ConvertFrom-ExcelColumnName ah
ConvertFrom-ExcelColumnName xfd
