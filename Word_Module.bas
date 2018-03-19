Attribute VB_Name = "Word_Module"
Sub ChangePlainDocumentToAMergeFields()
'Starting at the left of the word, it highlights word, unhighlights crlf, then changes into a mergefield
For X = 1 To 200
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend 'highlight line
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend    'dont select crlf
    If Selection.FormattedText <> "" And InStr(Selection.FormattedText, " ") = 0 Then   'FormattedText handles crlfs better than .Text
        ChangeSelectionToMergeField
    End If
    Selection.MoveRight Unit:=wdCharacter, Count:=1 'press right key
Next X
End Sub

Sub ChangeColumnLinesToMergeFields()
'Starting at the left of the word, it highlights word in a table row, then changes into a mergefield
For X = 1 To 200
    Selection.HomeKey Unit:=wdLine 'Home key
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend 'highlight line
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend    'dont select crlf
    If Selection.FormattedText <> "" And InStr(Selection.FormattedText, " ") = 0 And Asc(Selection) <> 160 Then   'FormattedText handles crlfs better than .Text. 160=table row
        ChangeSelectionToMergeField
    End If
    Selection.MoveDown Unit:=wdLine, Count:=1 'press down key
Next X
End Sub

Sub ChangeSelectionToMergeField()
Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:="MERGEFIELD  " + Selection.FormattedText + " ", PreserveFormatting:=True
End Sub

Public Sub ASCIICharacterCode()
    Application.StatusBar = "Character code: " & Asc(Selection)
End Sub

Sub PicResize(PercentSize As Integer)
     If Selection.InlineShapes.Count > 0 Then
         Selection.InlineShapes(1).ScaleHeight = PercentSize
         Selection.InlineShapes(1).ScaleWidth = PercentSize
     Else
         Selection.ShapeRange.ScaleHeight Factor:=(PercentSize / 100), _
           RelativeToOriginalSize:=msoCTrue
         Selection.ShapeRange.ScaleWidth Factor:=(PercentSize / 100), _
           RelativeToOriginalSize:=msoCTrue
     End If
 End Sub

Sub ResizeAllImagesInDocument()
    Dim PercentToSizeTo As Integer
    PercentToSizeTo = 25
    
    For Each oshp In ActiveDocument.Shapes
        oSho.Select
        PicResize (PercentToSizeTo)
    Next
    
    For Each oILShp In ActiveDocument.InlineShapes
        oILShp.Select
        PicResize (PercentToSizeTo)
    Next
End Sub



Sub remove_footer()
    Dim oHF As HeaderFooter
    Dim oSection As Section
     
    For Each oSection In ActiveDocument.Sections
        For Each oHF In oSection.Headers
            oHF.Range.Delete
        Next
        For Each oHF In oSection.Footers
            oHF.Range.Delete
        Next
    Next
End Sub

Sub change_all_pictures_to_greyscale()
    For Each oILShp In ActiveDocument.InlineShapes
        oILShp.PictureFormat.ColorType = msoPictureGrayscale
    Next
End Sub

Sub change_all_pictures_from_greyscale_to_colour()
    For Each oILShp In ActiveDocument.InlineShapes
        oILShp.PictureFormat.ColorType = msoPictureAutomatic
    Next
End Sub


Sub change_header_footer_pictures_to_greyscale()
    Dim sec As Section
    Dim sh As InlineShape
    
    For Each sec In ActiveDocument.Sections
        For Each sh In sec.Headers(wdHeaderFooterEvenPages).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureGrayscale
        Next sh
        For Each sh In sec.Footers(wdHeaderFooterEvenPages).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureGrayscale
        Next sh
        For Each sh In sec.Headers(wdHeaderFooterFirstPage).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureGrayscale
        Next sh
        For Each sh In sec.Footers(wdHeaderFooterFirstPage).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureGrayscale
        Next sh
        For Each sh In sec.Headers(wdHeaderFooterPrimary).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureGrayscale
        Next sh
        For Each sh In sec.Footers(wdHeaderFooterPrimary).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureGrayscale
        Next sh
    Next sec
End Sub

Sub change_header_footer_pictures_from_greyscale_to_colour()
    Dim sec As Section
    Dim sh As InlineShape
    
    For Each sec In ActiveDocument.Sections
        For Each sh In sec.Headers(wdHeaderFooterEvenPages).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureAutomatic
        Next sh
        For Each sh In sec.Footers(wdHeaderFooterEvenPages).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureAutomatic
        Next sh
        For Each sh In sec.Headers(wdHeaderFooterFirstPage).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureAutomatic
        Next sh
        For Each sh In sec.Footers(wdHeaderFooterFirstPage).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureAutomatic
        Next sh
        For Each sh In sec.Headers(wdHeaderFooterPrimary).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureAutomatic
        Next sh
        For Each sh In sec.Footers(wdHeaderFooterPrimary).Range.InlineShapes
            sh.PictureFormat.ColorType = msoPictureAutomatic
        Next sh
    Next sec
End Sub

Sub Add_pics_into_table_rows()
    'Useful for creating QRG / training documents
    Dim oTbl As Table, i As Long, j As Long, k As Long, StrTxt As String
     'Select and insert the Pics
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select image files and click OK"
        .Filters.Add "Images", "*.gif; *.jpg; *.jpeg; *.bmp; *.tif; *.png"
        .FilterIndex = 2
        If .Show = -1 Then
             'Add a 2-row by 2-column table with 7cm columns to take the images
            Set oTbl = Selection.Tables.Add(Selection.Range, 2, 2)
            Selection.MoveRight Unit:=wdCell
            With oTbl
                .AutoFitBehavior wdAutoFitWindow
                .Columns.Width = CentimetersToPoints(7)
                 .Style = "Table Grid"
            End With
            'Insert the Picture
            For i = 1 To .SelectedItems.Count
                Selection.MoveRight Unit:=wdCell
                Selection.MoveRight Unit:=wdCell
                ActiveDocument.InlineShapes.AddPicture _
                FileName:=.SelectedItems(i), LinkToFile:=False, SaveWithDocument:=True
            Next
        Else
        End If
    End With
End Sub

Sub SelectCurrentTable()
    Dim currentTableIndex As Integer
    
    ' Collapse the range to start so as to not have to deal with
    ' multi-segment ranges. Then check to make sure cursor is
    ' within a table.
    Selection.Collapse Direction:=wdCollapseStart
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Can only run this within a table"
        Exit Sub
    End If
    
    currentTableIndex = ActiveDocument.Range(0, Selection.Tables(1).Range.End).Tables.Count
    ActiveDocument.Tables(currentTableIndex).Select
End Sub

