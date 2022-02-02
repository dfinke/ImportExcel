# Add-ExcelImage Example

Adding pictures to an Excel worksheet is possible by calling the `AddPicture(name, image)`
method on the `Drawings` property of an `ExcelWorksheet` object.

The `Add-ExcelImage` example here demonstrates how to add a picture at a given
cell location, and optionally resize the row and column to fit the image.

Care has been taken in this example to get the image placement to be just inside
the cell border, and if the `-ResizeCell` switch is present, the height and width
of the row and column will be increased, if needed, so that the bottom right of
the image also lands just inside the cell border.

The Excel row and column sizes are measured in "point" units rather than pixels,
and at the moment a fixed multiplication factor is used to convert the size of
the image in pixels, to the corresponding height and width values in Excel.

You may find that images with a different DPI, or different resolution, DPI, or
text scaling options result in imperfect row and column sizing. 