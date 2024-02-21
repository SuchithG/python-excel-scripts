Sub UploadData()
    If AllDataValid Then
        Dim sourceWs As Worksheet
        Dim targetWb As Workbook
        Dim targetWs As Worksheet
        Dim lastRow As Long, lastTargetRow As Long
        Dim filePath As String

        ' Define the path to the target workbook
        filePath = "C:\Path\To\suchith-summary-stats.xlsx" ' Update this path to your target file's location

        ' Set the source worksheet to the active sheet
        Set sourceWs = ThisWorkbook.ActiveSheet
        
        ' Open the target workbook
        On Error Resume Next
        Set targetWb = Workbooks.Open(filePath)
        If Err.Number <> 0 Then
            MsgBox "Failed to open the target workbook. Check the file path.", vbCritical
            Exit Sub
        End If
        On Error GoTo 0

        ' Set the target worksheet, adjust "Sheet1" as necessary
        Set targetWs = targetWb.Sheets("Sheet1")
        
        ' Find the last row with data in both worksheets
        lastRow = sourceWs.Cells(sourceWs.Rows.Count, "A").End(xlUp).Row
        lastTargetRow = targetWs.Cells(targetWs.Rows.Count, "A").End(xlUp).Row + 1
        
        ' Copy data from source to target
        sourceWs.Range("A2:Z" & lastRow).Copy Destination:=targetWs.Range("A" & lastTargetRow)

        ' Save and close the target workbook
        targetWb.Save
        targetWb.Close SaveChanges:=True

        MsgBox "Data uploaded successfully.", vbInformation
    Else
        MsgBox "Data validation failed. Please validate the data before uploading.", vbExclamation
    End If
End Sub
