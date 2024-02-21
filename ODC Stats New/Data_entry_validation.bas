Sub PopulateResourceNameFromSheetName()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim resourceName As String
    Dim i As Long

    ' Set the worksheet to work with
    Set ws = ActiveSheet ' Or explicitly name your sheet like ThisWorkbook.Sheets("Richa")

    ' Get the active sheet name
    resourceName = ws.Name

    ' Find the last row with data in column A (assuming data starts in column A)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Populate the "Resource name" column for each row
    ' Adjust the column index as per your "Resource name" column's position
    For i = 2 To lastRow ' Start from row 2 to skip the header
        ws.Cells(i, "B").Value = resourceName ' Change "B" to your "Resource name" column letter
    Next i
End Sub
