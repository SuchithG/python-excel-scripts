Sub UploadDataToSummaryStats()
    Dim wbSource As Workbook, wbDest As Workbook
    Dim wsSuchith As Worksheet, wsCallsData As Worksheet
    Dim wsDestSuchith As Worksheet, wsDestCallsData As Worksheet
    Dim lastRowSourceSuchith As Long, lastRowSourceCallsData As Long
    Dim row As Long, sumCheck As Double
    Dim col As Integer, copyRangeSuchith As Range, copyRangeCallsData As Range
    Dim sourceFilePath As String, destFilePath As String
    
    On Error GoTo ErrorHandler
    
    ' Define full file paths
    sourceFilePath = "C:\Path\To\Your\Source\Workbook\Suchith Daily Stats.xlsm"
    destFilePath = "C:\Path\To\Your\Destination\Workbook\ODC Summary Stats\suchith-summary-stats.xlsx"
    
    ' Open source workbook
    Set wbSource = Workbooks.Open(sourceFilePath)
    Set wsSuchith = wbSource.Sheets("Suchith")
    Set wsCallsData = wbSource.Sheets("Calls Data")
    
    ' Find last row with data in source worksheets
    lastRowSourceSuchith = wsSuchith.Cells(wsSuchith.Rows.Count, "A").End(xlUp).Row
    lastRowSourceCallsData = wsCallsData.Cells(wsCallsData.Rows.Count, "A").End(xlUp).Row
    
    ' Debug statements
    Debug.Print "Last row in 'Suchith': " & lastRowSourceSuchith
    Debug.Print "Last row in 'Calls Data': " & lastRowSourceCallsData
    
    ' Validate data in 'Suchith' sheet and check for data presence
    If lastRowSourceSuchith > 1 Then
        For row = 2 To lastRowSourceSuchith ' Assuming data starts from row 2
            sumCheck = Application.WorksheetFunction.Sum(wsSuchith.Range("I" & row & ":N" & row))
            If sumCheck < 1 Then
                MsgBox "Sum of columns I to N in row " & row & " of 'Suchith' sheet must be greater than or equal to 1.", vbExclamation
                GoTo CleanUp
            End If
            
            For col = 1 To 8 ' Columns A to H
                If IsEmpty(wsSuchith.Cells(row, col).Value) Then
                    MsgBox "Column " & Chr(64 + col) & " in row " & row & " of 'Suchith' sheet cannot be blank.", vbExclamation
                    GoTo CleanUp
                End If
            Next col
            For col = 15 To 17 ' Columns O to Q
                If IsEmpty(wsSuchith.Cells(row, col).Value) Then
                    MsgBox "Column " & Chr(64 + col) & " in row " & row & " of 'Suchith' sheet cannot be blank.", vbExclamation
                    GoTo CleanUp
                End If
            Next col
        Next row
    End If
    
    ' Open destination workbook
    Set wbDest = Workbooks.Open(destFilePath)
    Set wsDestSuchith = wbDest.Sheets("Suchith")
    Set wsDestCallsData = wbDest.Sheets("Calls Data")
    
    ' Copy data from source 'Suchith' sheet to destination if there are data rows
    If lastRowSourceSuchith > 1 Then
        Set copyRangeSuchith = wsSuchith.Range("A2:Q" & lastRowSourceSuchith)
        copyRangeSuchith.Copy wsDestSuchith.Cells(wsDestSuchith.Rows.Count, "A").End(xlUp).Offset(1)
        ' Clear contents of copied ranges in source worksheet
        wsSuchith.Range("A2:Q" & lastRowSourceSuchith).ClearContents
    End If
    
    ' Copy data from source 'Calls Data' sheet to destination if there are data rows
    If lastRowSourceCallsData > 1 Then
        Set copyRangeCallsData = wsCallsData.Range("A2:F" & lastRowSourceCallsData)
        copyRangeCallsData.Copy wsDestCallsData.Cells(wsDestCallsData.Rows.Count, "A").End(xlUp).Offset(1)
        ' Clear contents of copied ranges in source worksheet
        wsCallsData.Range("A2:F" & lastRowSourceCallsData).ClearContents
    End If
    
    ' Save and close destination workbook
    wbDest.Close SaveChanges:=True
    
    ' Close source workbook without saving changes
    wbSource.Close SaveChanges:=False
    
    MsgBox "Data uploaded successfully.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    GoTo CleanUp

CleanUp:
    ' Ensure workbooks are closed properly in case of error
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    If Not wbDest Is Nothing Then wbDest.Close SaveChanges:=False
End Sub
