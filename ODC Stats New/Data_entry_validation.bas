Global AllDataValid As Boolean

Sub ValidateData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow As Long, i As Long, colIndex As Long
    Dim isValidRow As Boolean
    Dim errorRows As String
    Dim numericEntryRequiredColumns As Variant
    numericEntryRequiredColumns = Array(11, 12, 13, 14, 15) ' Setup, Amend, Review, Closure, Exceptions
    
    AllDataValid = True
    errorRows = ""
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    For i = 2 To lastRow
        isValidRow = True ' Assume the row is valid initially
        
        ' Check PDF Name is not blank and not purely numeric
        If ws.Cells(i, 6).Value = "" Or IsNumeric(ws.Cells(i, 6).Value) Then
            isValidRow = False
        End If
        
        ' Check if at least one of the specific columns has a numeric entry greater than 0
        Dim hasPositiveNumeric As Boolean
        hasPositiveNumeric = False
        For Each colIndex In numericEntryRequiredColumns
            If ws.Cells(i, colIndex).Value > 0 Then
                hasPositiveNumeric = True
                Exit For
            End If
        Next
        
        If Not hasPositiveNumeric Then isValidRow = False ' No positive numeric entries found
        
        ' If there is an entry in "PDF name", there cannot be an entry in "Activity" column
        If ws.Cells(i, 9).Value <> "" Then isValidRow = False ' Activity has an entry
        
        ' Mandatory fields: Resource Name, Date, Month, Region, 2 eye, 4 eye must be filled
        For Each colIndex In Array(2, 3, 4, 5, 16, 17) ' Checking mandatory fields
            If IsEmpty(ws.Cells(i, colIndex).Value) Then
                isValidRow = False
                Exit For
            End If
        Next
        
        ' Coloring rows based on validation
        If isValidRow Then
            ws.Rows(i).Interior.Color = RGB(144, 238, 144) ' Light green for valid rows
        Else
            ws.Rows(i).Interior.Color = RGB(255, 99, 71) ' Light red for invalid rows
            AllDataValid = False
            If Len(errorRows) > 0 Then errorRows = errorRows & ", "
            errorRows = errorRows & i
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    If AllDataValid Then
        MsgBox "All data is valid. Ready to upload.", vbInformation
    Else
        MsgBox "Validation failed on rows: " & errorRows, vbCritical
    End If
End Sub
