Sub UploadDataToSummaryStats()
    Dim wbSource As Workbook, wbDest As Workbook
    Dim wsSuchith As Worksheet, wsCallsData As Worksheet
    Dim wsDestSuchith As Worksheet, wsDestCallsData As Worksheet
    Dim lastRowSourceSuchith As Long, lastRowSourceCallsData As Long
    Dim copyRangeSuchith As Range, copyRangeCallsData As Range
    Dim sourceFilePath As String, destFilePath As String
    Dim row As Long, sumCheck As Double
    Dim rowData As Range
    Dim response As VbMsgBoxResult

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

    ' Open destination workbook
    Set wbDest = Workbooks.Open(destFilePath)
    Set wsDestSuchith = wbDest.Sheets("Suchith")
    Set wsDestCallsData = wbDest.Sheets("Calls Data")

    ' Copy data from source 'Suchith' sheet to destination if there are data rows
    If lastRowSourceSuchith > 1 Then
        For row = 2 To lastRowSourceSuchith
            ' Check if all mandatory values are present and sum of values from "Count" column to "Exceptions" column is greater than 0
            If Application.CountA(wsSuchith.Range("A" & row & ":H" & row)) = 8 And Application.CountA(wsSuchith.Range("O" & row & ":Q" & row)) = 3 Then
                sumCheck = Application.WorksheetFunction.Sum(wsSuchith.Range("I" & row & ":N" & row))
                If sumCheck > 0 Then
                    Set copyRangeSuchith = wsSuchith.Range("A" & row & ":Q" & row)
                    copyRangeSuchith.Copy wsDestSuchith.Cells(wsDestSuchith.Rows.Count, "A").End(xlUp).Offset(1)
                    ' Clear contents of copied range in source worksheet
                    wsSuchith.Rows(row).ClearContents
                Else
                    response = MsgBox("Sum of values from 'Count' to 'Exceptions' in row " & row & " is not greater than 0. Do you want to discard this row?", vbYesNo + vbExclamation)
                    If response = vbYes Then
                        ' Skip copying and clear the row in the source sheet
                        wsSuchith.Rows(row).ClearContents
                    Else
                        ' Do not close the workbook, allow user to make changes
                        GoTo CleanupSkipClose
                    End If
                End If
            Else
                response = MsgBox("Mandatory values are missing in row " & row & ". Do you want to discard this row?", vbYesNo + vbExclamation)
                If response = vbYes Then
                    ' Skip copying and clear the row in the source sheet
                    wsSuchith.Rows(row).ClearContents
                Else
                    ' Do not close the workbook, allow user to make changes
                    GoTo CleanupSkipClose
                End If
            End If
        Next row
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
    GoTo CleanupSkipClose

CleanupSkipClose:
    ' Ensure destination workbook is closed properly in case of error
    If Not wbDest Is Nothing Then wbDest.Close SaveChanges:=False
    ' Exit without closing the source workbook
    Exit Sub
End Sub
