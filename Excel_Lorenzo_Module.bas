Attribute VB_Name = "Lorenzo_Module"
Sub Remove_Inpatient_Headers()
Attribute Remove_Inpatient_Headers.VB_ProcData.VB_Invoke_Func = "h\n14"
'
' Remove_Inpatient_Headers Macro
'
' Keyboard Shortcut: Ctrl+h
'
    Application.Goto Reference:="R12C1"
    ActiveCell.FormulaR1C1 = "=ISNUMBER(VALUE(RC[2]))"
    ActiveCell.Offset(0, 0).Range("A1").Select
    Selection.Copy
    ActiveCell.SpecialCells(xlLastCell).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    ActiveCell.Offset(-2, 0).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.AutoFilter
    Selection.AutoFilter Field:=1, Criteria1:="FALSE"
    Application.Goto Reference:="R12C1"
    ActiveCell.Offset(1, 0).Range("A1").Select
    SelectNextVisibleCell
    Range(Selection, Selection.End(xlDown)).Select
    Application.DisplayAlerts = False   '"This operation will cause some merged cells to unmerge"
    Selection.EntireRow.Delete
    Application.DisplayAlerts = True
    ActiveSheet.ShowAllData
    Application.Goto Reference:="R12C1"
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A3").Select
    'Bed
    Columns("J:J").ColumnWidth = 7
    'Specialty
    Columns("Q:Q").ColumnWidth = 7
End Sub
Sub SelectNextVisibleCell()
ActiveCell.Offset(1, 0).Select
Do Until ActiveCell.Height <> 0
    ActiveCell.Offset(1, 0).Select
Loop
End Sub
