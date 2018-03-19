Attribute VB_Name = "Powerpoint_Module"
Sub Resize()
    On Error GoTo 1
    
    With ActiveWindow.Selection.ShapeRange
        '.Height = 900
        .Width = 700
        .Left = 10
        .Top = 0
        '.ZOrder msoSendToBack  'This sends picture to the back
    End With
1:
End Sub

