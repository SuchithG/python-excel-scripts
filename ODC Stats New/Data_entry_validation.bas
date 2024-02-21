Global AllDataValid As Boolean

Sub ValidateData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow As Long, i As Long, j As Long, colIndex As Long
    Dim isValidRow As Boolean, hasNumericEntry As Boolean, errorRows As String
    
    AllDataValid = True
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    errorRows = ""
    
    Application.ScreenUpdating = False
    
    For i = 2 To lastRow ' Assuming data starts at row 2
        isValidRow = True
        hasNumericEntry = False
        
        ' PDF Name validation and numeric entry check, assuming "PDF Name" is in column E (5)
        If ws.Cells(i, 5).Value = "" Or IsNumeric(ws.Cells(i, 5).Value) Then
            isValidRow = False
        Else
            ' Check for numeric entry > 0 in specified columns
            Dim checkColumns As Variant
            checkColumns = Array(10, 11, 12, 13, 14) ' Columns for Setup, Amend, Review, Closure, Exceptions
            
            For j = LBound(checkColumns) To UBound(checkColumns)
                colIndex = checkColumns(j)
                If IsNumeric(ws.Cells(i, colIndex).Value) And ws.Cells(i, colIndex).Value > 0 Then
                    hasNumericEntry = True
                    Exit For
                End If
            Next j
            
            ' If no valid numeric entry found or Activity column is filled
            If Not hasNumericEntry Or ws.Cells(i, 16).Value <> "" Then
                isValidRow = False
            End If
        End If
        
        ' Mandatory fields check
        mandatoryColumns = Array(2, 3, 4, 5, 6, 7) ' Replace with actual column indexes for mandatory fields
        
        For j = LBound(mandatoryColumns) To UBound(mandatoryColumns)
            colIndex = mandatoryColumns(j)
            If IsEmpty(ws.Cells(i, colIndex).Value) Then
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
