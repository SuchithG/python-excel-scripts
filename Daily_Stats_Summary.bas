Sub progQ()
    If IsEmpty(Range("J2")) Then
        MsgBox "Select Date from given drop down in Cell: J2"
    Else
        upload
    End If
End Sub

Sub upload()
    'Unprotect a worksheet with a password
    Sheets("Macro Template").Unprotect Password:="A" 'Unprotect a worksheet with a password
    Sheets("Macro Template").Unprotect Password:="A"
    '
    Columns("AF").Select
    Range("F1").Activate
    Selection.EntireColumn.Hidden = False
    Range("A3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$3:$Q$35").AutoFilter Field:=1, Criteria1:="Add"
    Rows("4:57").Select
    Selection.Copy
    Workbooks.Open Filename:= _
        "filepath\Test Daily Stats Summary\xlsm.xlsx"
    Application.Goto Reference:="R7500C1"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").Select
    Range("A74962").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    ActiveWindow.SmallScroll Down:=-33
    Range("A1").Select
    ActiveWorkbook.Save
    ActiveWindow.Close
    Selection.Copy
    Sheets("Reference Data").Select
    Application.Goto Reference:="R7500C1"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").Select
    Range("A74962").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    ActiveWindow.SmallScroll Down:=-33
    Range("A1").Select
    ActiveWorkbook.Save
    ActiveWindow.Close
    Selection.Copy
    Sheets("Reference Data").Select
    Application.Goto Reference:="R7500C1"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").Select
    Range("A74962").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    Range("C37").Select
    ActiveWindow.SmallScroll Down:=-60
    Range("A1").Select
    Sheets("Macro Template").Select
    Selection.AutoFilter
    Columns("A:E").Select
    Range("E1").Activate
    Selection.EntireColumn.Hidden = True
    Range("F4:R35").Select
    Selection.ClearContents
    Range("J2").Select
    Selection.ClearContents
    Range("K13").Select
    Sheets("Reference Data").Select
    Sheets("Macro Template").Select
    'Protect worksheet with a password
    Sheets("Macro Template").Protect Password:="A"
    MsgBox "Stats uploaded. To validate, Check in Sheet = Reference Data in your stats file"
    End Sub

    Sub deleQ()
        Range("S2").Select
        ActiveCell.FormulaR1C1 = "=TODAY()-RC[-1]"
        Columns("S:S").Select
        Selection.NumberFormat = "[<=409]d-mmm-yy;@"
        Selection.NumberFormat = "General"
        Range("S2").Select
        Selection.AutoFill Destination:=Range("S2:S158")
        Range("S2:S158").Select
        Range("T2").Select
        ActiveCell.FormulaR1C1 = "=IF(RC[-1]>25,""Del"",""Keep"")"
        Columns("S:T").Select
        With Selection.Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
        Range("T2").Select
        ActiveCell.FormulaR1C1 = "=IF(RC[-1]>25,""Del"",""Keep"")"
        Range("T2").Select
        Selection.AutoFill Destination:=Range("T2:T158")
        Range("T2:T158").Select
        Range("S1").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$T$158").AutoFilter Field:=20, Criteria1:="Del"
        Rows("2:2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Rows("2:257").Select
        Selection.Delete Shift:=xlUp
        Range("G44").Select
        Selection.AutoFilter
        Columns("S:T").Select
        Selection.ClearContents
        Range("J26").Select
        Sheets("Macro Template").Select
        ActiveWindow.SmallScroll Down:=-12
    End Sub