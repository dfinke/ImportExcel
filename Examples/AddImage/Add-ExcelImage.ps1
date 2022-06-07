function Add-ExcelImage {
    <#
    .SYNOPSIS
        Adds an image to a worksheet in an Excel package.
    .DESCRIPTION
        Adds an image to a worksheet in an Excel package using the
        `WorkSheet.Drawings.AddPicture(name, image)` method, and places the
        image at the location specified by the Row and Column parameters.
        
        Additional position adjustment can be made by providing RowOffset and
        ColumnOffset values in pixels.
    .EXAMPLE
        $image = [System.Drawing.Image]::FromFile($octocat)
        $xlpkg = $data | Export-Excel -Path $path -PassThru
        $xlpkg.Sheet1 | Add-ExcelImage -Image $image -Row 4 -Column 6 -ResizeCell
        
        Where $octocat is a path to an image file, and $data is a collection of
        data to be exported, and $path is the output path for the Excel document,
        Add-Excel places the image at row 4 and column 6, resizing the column
        and row as needed to fit the image.
    .INPUTS
        [OfficeOpenXml.ExcelWorksheet]
    .OUTPUTS
        None
    #>
    [CmdletBinding()]
    param(
        # Specifies the worksheet to add the image to.
        [Parameter(Mandatory, ValueFromPipeline)]
        [OfficeOpenXml.ExcelWorksheet]
        $WorkSheet,

        # Specifies the Image to be added to the worksheet.
        [Parameter(Mandatory)]
        [System.Drawing.Image]
        $Image,        

        # Specifies the row where the image will be placed. Rows are counted from 1.
        [Parameter(Mandatory)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $Row,

        # Specifies the column where the image will be placed. Columns are counted from 1.
        [Parameter(Mandatory)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $Column,

        # Specifies the name to associate with the image. Names must be unique per sheet.
        # Omit the name and a GUID will be used instead.
        [Parameter()]
        [string]
        $Name,

        # Specifies the number of pixels to offset the image on the Y-axis. A
        # positive number moves the image down by the specified number of pixels
        # from the top border of the cell.
        [Parameter()]
        [int]
        $RowOffset = 1,

        # Specifies the number of pixels to offset the image on the X-axis. A
        # positive number moves the image to the right by the specified number
        # of pixels from the left border of the cell.
        [Parameter()]
        [int]
        $ColumnOffset = 1,

        # Increase the column width and row height to fit the image if the current
        # dimensions are smaller than the image provided.
        [Parameter()]
        [switch]
        $ResizeCell
    )

    begin {
        if ($IsWindows -eq $false) {
            throw "This only works on Windows and won't run on $([environment]::OSVersion)"
        }
        
        <#
          These ratios work on my machine but it feels fragile. Need to better
          understand how row and column sizing works in Excel and what the
          width and height units represent.
        #>
        $widthFactor = 1 / 7
        $heightFactor = 3 / 4
    }

    process {
        if ([string]::IsNullOrWhiteSpace($Name)) {
            $Name = (New-Guid).ToString()
        }
        if ($null -ne $WorkSheet.Drawings[$Name]) {
            Write-Error "A picture with the name `"$Name`" already exists in worksheet $($WorkSheet.Name)."
            return
        }

        <#
          The row and column offsets of 1 ensures that the image lands just
          inside the gray cell borders at the top left.
        #>
        $picture = $WorkSheet.Drawings.AddPicture($Name, $Image)
        $picture.SetPosition($Row - 1, $RowOffset, $Column - 1, $ColumnOffset)
        
        if ($ResizeCell) {
            <#
              Adding 1 to the image height and width ensures that when the
              row and column are resized, the bottom right of the image lands
              just inside the gray cell borders at the bottom right.
            #>
            $width = $widthFactor * ($Image.Width + 1)
            $height = $heightFactor * ($Image.Height + 1)
            $WorkSheet.Column($Column).Width = [Math]::Max($width, $WorkSheet.Column($Column).Width)
            $WorkSheet.Row($Row).Height = [Math]::Max($height, $WorkSheet.Row($Row).Height)
        }
    }
}