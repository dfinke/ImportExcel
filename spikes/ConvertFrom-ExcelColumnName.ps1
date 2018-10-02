function ConvertFrom-ExcelColumnName {
    param($columnName)

    $sum=0
    $columnName.ToCharArray() |
        ForEach-Object {
            $sum*=26
            $sum+=[char]$_.tostring().toupper()-[char]'A'+1
        }
    $sum
}

ConvertFrom-ExcelColumnName -columnName p
ConvertFrom-ExcelColumnName -columnName ah
ConvertFrom-ExcelColumnName -columnName xfd
