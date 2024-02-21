Global AllDataValid As Boolean

Sub ValidateData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow As Long, i As Long, j As Long
    Dim isValidRow As Boolean, hasNumericEntry As Boolean, errorRows As String
    
    AllDataValid = True
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    errorRows = ""
    
    Application.ScreenUpdating = False
    
    For i = 2 To lastRow ' Assuming data starts at row 2
        isValidRow = True
        hasNumericEntry = False
        
        ' PDF Name validation (Column E = 5)
        If ws.Cells(i, 6).Value = "" Or IsNumeric(ws.Cells(i, 6).Value) Then
            isValidRow = False
        Else
            ' Check for numeric entry > 0 in specified columns (Setup = 11, Amend = 12, Review = 13, Closure = 14, Exceptions = 15)
            For Each j In Array(11, 12, 13, 14, 15)
                If IsNumeric(ws.Cells(i, j).Value) And ws.Cells(i, j).Value > 0 Then
                    hasNumericEntry = True
                    Exit For
                End If
            Next j
            
            ' If no valid numeric entry found or Activity column (Column J = 10) is filled
            If Not hasNumericEntry Or ws.Cells(i, 9).Value <> "" Then ' Corrected based on column order
                isValidRow = False
            End If
        End If
        
        ' Mandatory fields check (Resource Name = 2, Date = 3, Month = 4, Region = 5, 2 eye = 16, 4 eye = 17, Actual Date of upload = 18)
        For Each j In Array(2, 3, 4, 5, 16, 17, 18)
            If IsEmpty(ws.Cells(i, j).Value) Then
                isValidRow = False
                Exit For
            End If
        Next
        
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
