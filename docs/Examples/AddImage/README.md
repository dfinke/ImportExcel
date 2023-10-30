# Add-ExcelImage Example

Adding pictures to an Excel worksheet is possible by calling the `AddPicture(name, image)`
method on the `Drawings` property of an `ExcelWorksheet` object.

The `Add-ExcelImage` example here demonstrates how to add a picture at a given
cell location, and optionally resize the row and column to fit the image.

## Running the example

To try this example, run the script `AddImage.ps1`. The `Add-ExcelImage`
function will be dot-sourced, and an Excel document will be created in the same
folder with a sample data set. The Octocat image will then be embedded into
Sheet1.

The creation of the Excel document and the `System.Drawing.Image` object
representing Octocat are properly disposed within a `finally` block to ensure
that the resources are released, even if an error occurs in the `try` block.

## Note about column and row sizing

Care has been taken in this example to get the image placement to be just inside
the cell border, and if the `-ResizeCell` switch is present, the height and width
of the row and column will be increased, if needed, so that the bottom right of
the image also lands just inside the cell border.

The Excel row and column sizes are measured in "point" units rather than pixels,
and a fixed multiplication factor is used to convert the size of the image in
pixels, to the corresponding height and width values in Excel.

It's possible that different DPI or text scaling options could result in
imperfect column and row sizing and if a better strategy is found for converting
the image dimensions to column and row sizes, this example will be updated.
