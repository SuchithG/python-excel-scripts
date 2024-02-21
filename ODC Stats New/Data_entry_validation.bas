Private Sub Worksheet_Change(ByVal Target As Range)
    Dim resourceNameColumn As Integer: resourceNameColumn = 2 ' "Resource name" in column B
    Dim pdfNameColumn As Integer: pdfNameColumn = 5 ' "PDF Name" in column E
    Dim checkColumns As Variant: checkColumns = Array(10, 11, 12, 13, 14) ' Columns for Setup, Amend, Review, Closure, Exceptions
    Dim appColumn As Integer: appColumn = 15 ' "Application" in column O
    Dim activityColumn As Integer: activityColumn = 16 ' "Activity" in column P
    Dim ws As Worksheet: Set ws = Me
    Dim hasNumeric As Boolean: hasNumeric = False
    Dim i As Long

    On Error GoTo ErrorHandler
    Application.EnableEvents = False ' Prevent recursive event triggering

    ' Check for changes in specific columns to minimize unnecessary processing
    If Not Intersect(Target, ws.Columns(resourceNameColumn)) Is Nothing Or _
       Not Intersect(Target, ws.Columns(pdfNameColumn)) Is Nothing Or _
       Not Intersect(Target, ws.Columns(appColumn)) Is Nothing Or _
       Not Intersect(Target, ws.Columns(activityColumn)) Is Nothing Then
        
        ' Populate "Resource name" for any row change if the column is specified
        If Not Intersect(Target, ws.Columns(resourceNameColumn)) Is Nothing Then
            ws.Cells(Target.Row, resourceNameColumn).Value = ws.Name
        End If

        ' Validate "PDF Name" entries
        If Not Intersect(Target, ws.Columns(pdfNameColumn)) Is Nothing Then
            ' Check if PDF Name is not blank and not purely numeric
            If Not IsEmpty(Target.Value) And Not IsNumeric(Target.Value) Then
                ' Check for at least one numeric entry in specified columns
                For Each i In checkColumns
                    If IsNumeric(ws.Cells(Target.Row, i).Value) And ws.Cells(Target.Row, i).Value <> "" Then
                        hasNumeric = True
                        Exit For
                    End If
                Next i
                
                ' If no numeric value found, clear PDF Name and notify
                If Not hasNumeric Then
                    MsgBox "Please enter a numeric value in at least one of the specified columns (Setup, Amend, Review, Closure, Exceptions) if PDF Name is provided."
                    Target.ClearContents
                End If
                
                ' Ensure "Application" and "Activity" are blank if PDF Name is filled
                If hasNumeric Then
                    Dim response As VbMsgBoxResult
                    response = MsgBox("Clear 'Application' and 'Activity' fields for this row?", vbQuestion + vbYesNo, "Confirm")
                    If response = vbYes Then
                        ws.Cells(Target.Row, appColumn).ClearContents
                        ws.Cells(Target.Row, activityColumn).ClearContents
                    End If
                End If
            End If
        End If
    End If

    GoTo ExitHandler

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    GoTo ExitHandler

ExitHandler:
    Application.EnableEvents = True ' Re-enable events
End Sub
