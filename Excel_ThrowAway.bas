Attribute VB_Name = "ThrowAway"
Sub MoveToComplete()
Attribute MoveToComplete.VB_ProcData.VB_Invoke_Func = " \n14"
'move completed task to next sheet
    ActiveCell.Rows("1:1").EntireRow.Select
    ActiveCell.Activate
    Selection.Cut
    ActiveSheet.Next.Select
    ActiveCell.SpecialCells(xlLastCell).Select
    ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "g"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Previous.Select
    Selection.Delete Shift:=xlUp
    ActiveCell.Offset(0, 0).Range("A1").Select
End Sub
