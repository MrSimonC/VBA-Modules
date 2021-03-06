Attribute VB_Name = "VisioVBA"
Sub QuickTextToBottom()
Attribute QuickTextToBottom.VB_ProcData.VB_Invoke_Func = "q"
 
  '// 2007.11.07
  '// Visio Guy
  '//
  '// Adjusts the text block of selected shapes so that
  '// the text is at the bottom of the shape. This matches
  '// the default text position for inserted images.
 
  Dim sel As Visio.Selection
  Dim shp As Visio.Shape
  Set sel = Visio.ActiveWindow.Selection
 
  For Each shp In sel
 
    '// 'Add' the Text Transfomrm section, if it's not there:
    If Not (shp.RowExists(Visio.VisSectionIndices.visSectionObject, _
          Visio.VisRowIndices.visRowTextXForm, _
          Visio.VisExistsFlags.visExistsAnywhere)) Then
 
    Call shp.AddRow(Visio.VisSectionIndices.visSectionObject, _
          Visio.VisRowIndices.visRowTextXForm, _
          Visio.VisRowTags.visTagDefault)
 
    End If
 
    '// Set the text transform formulas:
    shp.CellsU("TxtHeight").FormulaForceU = "Height*0"
    shp.CellsU("TxtWidth").FormulaForceU = "Width*2"
    shp.CellsU("TxtPinY").FormulaForceU = "Height*0"
 
    '// Set the paragraph alignment formula:
    shp.CellsU("VerticalAlign").FormulaForceU = "0"
 
  Next shp
 
End Sub

