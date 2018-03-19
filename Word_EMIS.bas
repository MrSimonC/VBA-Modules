Attribute VB_Name = "EMIS"
Sub EMISConsultationToText()
Attribute EMISConsultationToText.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.simonTest"
    'v1.0 by Simon Crouch 23Aug17
    'Find Consultation Text
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "consultation text"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    
    'Select the current active outside table
    Dim currentTableIndex As Integer
    Selection.Collapse Direction:=wdCollapseStart
    If Not Selection.Information(wdWithInTable) Then
        Exit Sub
    End If
    currentTableIndex = ActiveDocument.Range(0, Selection.Tables(1).Range.End).Tables.Count
    ActiveDocument.Tables(currentTableIndex).Select
    
    'Convert
    Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=True

    'Replace tabs out
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    With Selection.Find
        .Text = "^t^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop 'wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop 'wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop 'wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop 'wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop 'wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    With Selection.Find
        .Text = "^t^t"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindStop 'wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t^t"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindStop 'wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop 'wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'Justify
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
End Sub

