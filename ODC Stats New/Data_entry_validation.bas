Global AllDataValid As Boolean

Sub ValidateData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim isValidRow As Boolean
    Dim hasNumericEntry As Boolean
    Dim errorRows As String
    
    AllDataValid = True
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    errorRows = ""
    
    Application.ScreenUpdating = False
    
    For i = 2 To lastRow ' Assuming data starts at row 2
        isValidRow = True
        hasNumericEntry = False
        
        ' Rule 1 & 2: PDF Name validation and numeric entry check
        If ws.Cells(i, 5).Value = "" Or Not IsNumeric(ws.Cells(i, 5).Value) Then ' Assuming "PDF Name" is in column E
            isValidRow = False
        Else
            ' Check for numeric entry > 0 in specified columns
            For Each j In Array(10, 11, 12, 13, 14) ' Columns for Setup, Amend, Review, Closure, Exceptions
                If ws.Cells(i, j).Value > 0 Then
                    hasNumericEntry = True
                    Exit For
                End If
            Next j
            
            ' If no valid numeric entry found or Activity column is filled
            If Not hasNumericEntry Or ws.Cells(i, "ActivityColumnIndex").Value <> "" Then
                isValidRow = False
            End If
        End If
        
        ' Rule 3: Mandatory fields check
        For Each j In Array(2, "DateColumnIndex", "MonthColumnIndex", "RegionColumnIndex", "2EyeColumnIndex", "4EyeColumnIndex") ' Adjust column indexes
            If IsEmpty(ws.Cells(i, j).Value) Then
                isValidRow = False
                Exit For
            End If
        Next j
        
        ' Apply color coding based on validation
        If isValidRow Then
            ws.Rows(i).Interior.Color = RGB(144, 238, 144) ' Light green for valid rows
        Else
            ws.Rows(i).Interior.Color = RGB(255, 99, 71) ' Light red for invalid rows
            AllDataValid = False
            errorRows = errorRows & i & ", "
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    If AllDataValid Then
        MsgBox "All data is valid. Ready to upload.", vbInformation
    Else
        MsgBox "Some rows have errors: " & Left(errorRows, Len(errorRows) - 2), vbCritical
    End If
End Sub
