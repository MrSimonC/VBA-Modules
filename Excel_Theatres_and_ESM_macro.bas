Attribute VB_Name = "Theatres_and_ESM_macro"
Sub ESM_to_CSV()
DCW_to_CSV_file "ESM", "Default Schedules - Templates", "Modified?", "C:\Cerner Audit\ESMds.csv", "C:\Cerner Audit\ESM for simon audit.xlsx", "C:\Cerner Audit"
End Sub

Sub Theatres_to_CSV()
DCW_to_CSV_file "Theatres", "Session Times", "Template Build Name", "C:\Cerner Audit\TheatresDCW.csv", "", "I:\Cerner\Project Deliverables\DCWs\DCWs\DCWs\Theatres\New Hospital"
End Sub

Sub DCW_to_CSV_file(FileSearch As String, DCWtab As String, DCWSearch As String, CSVDefaultFilePath As String, DCWDefaultFilePath As String, DCWSearchPath As String)
' Highlights some data, copies to another workbook, formats, then saves as CSV
' Assumes:
    ' Your data is between top right(DCWSearch) and first bottom entry of column A
    ' There is only 1 cell with DCWSearch contents

    Dim DCWFile As String       'workbook name of dcw
    Dim DCWFilePath As String   'Path of dcw
    Dim CSVFile As String       'workbook name of csv
    Dim CSVFilePath As String   'Path of csv
        
    'Get DCW File
    If FileFolderExists(DCWDefaultFilePath) Then
        DCWFilePath = DCWDefaultFilePath
        If GetFileName(DCWFilePath, DCWFile) = False Then Exit Sub
    Else
        If Not ChooseFile("Please choose the " & FileSearch & " DCW", DCWSearchPath, DCWFilePath) _
            Or Not FileFolderExists(DCWFilePath) _
            Or GetFileName(DCWFilePath, DCWFile) = False _
            Or InStr(DCWFile, FileSearch) = 0 Then
            MsgBox "This doesn't look like the '" & FileSearch & "' DCW.", vbCritical
            Exit Sub
        End If
    End If

    'Check CSV File
    If FileFolderExists(CSVDefaultFilePath) Then
        CSVFilePath = CSVDefaultFilePath
        If GetFileName(CSVDefaultFilePath, CSVFile) = False Then Exit Sub
    Else
        If Not ChooseFile("Please choose the CSV output file", "C:\Cerner Audit\", CSVFilePath) _
        Or Not FileFolderExists(CSVFilePath) _
        Or GetFileName(CSVFilePath, CSVFile) = False _
        Or InStr(CSVFile, FileSearch) = 0 Then
            MsgBox "This doesn't look like the " & FileSearch & " CSV output file.", vbCritical
                Exit Sub
        End If
    End If
    
    'Check paths are different
    If CSVFilePath = DCWFilePath Then
        MsgBox "You can't choose the same file as input and output.", vbCritical, "Error: Can't input and output to same file"
        Exit Sub
    End If
    
    'Open Files
    Application.ScreenUpdating = False
    If CheckFileIsOpen(DCWFile) = False Then Workbooks.Open DCWFilePath, ReadOnly:=True
        'Check Sheet exists, i.e. that you're in the ESM DCW
        Windows(DCWFile).Activate
        If SheetExists(DCWtab) = False Then
            MsgBox "I can't see the " + DCWtab + " tab on the " + FileSearch + " DCW. Is it there or did you choose the right file?", vbCritical, "Error: Can't see DCW"
            Workbooks(DCWFile).Close SaveChanges:=False
            Application.ScreenUpdating = True
            Exit Sub
        End If
    If CheckFileIsOpen(CSVFile) = False Then Workbooks.Open CSVFilePath
        Windows(CSVFile).Activate
        If SheetExists(DCWtab) Then
            MsgBox "This CSV file looks like the " + FileSearch + " DCW as I can find a " + DCWtab + " tab. For saftey, please try this macro again, ensuring the csv file is correctly chosen and doesn't have a " + DCWtab + " tab.", vbCritical, "Error: CSV file looks like DCW file"
            Workbooks(DCWFile).Close SaveChanges:=False
            Workbooks(CSVFile).Close SaveChanges:=False
            Application.ScreenUpdating = True
            Exit Sub
        End If

    'Start the work
    Windows(CSVFile).Activate 'Blank the ESMds.csv file
        Cells.Select
        Selection.ClearContents
        Reset_Used_Range    'important, otherwise you get a trailing "'" on the end of each row
        Range("A1").Select
    Windows(DCWFile).Activate   'Copy the Data from DCW
        Sheets(DCWtab).Select
        If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False    'Completely remove filters
        Range("A1").Select  'Find (DCWSearch) in entire sheet, then Ctrl+Shift down and left
        Find_and_Select_Range DCWSearch, "A" 'search current row for DCWSearch, mark as top right, then down-up of column A is bottom-left
        Selection.Copy
    Windows(CSVFile).Activate 'Paste in CSV file
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    Insert_an_id_Row    'used for SQL - to autopopulate a RowID in SQL
    
    If FileSearch = "ESM" And DCWtab = "Default Schedules - Templates" Then ESM_DCW_DefaultSchedules_clean_up   'Run Clean Up. Necessary for date when saving as csv
    If FileSearch = "Theatres" And DCWtab = "Session Times" Then Theatres_DCW_clean_up   'Run Clean Up. Necessary for date when saving as csv

    Workbooks(CSVFile).Close SaveChanges:=True
    Workbooks(DCWFile).Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    MsgBox "CSV file processed successfully."
End Sub

Sub Insert_an_id_Row()
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "id"
End Sub

Sub ESM_DCW_DefaultSchedules_clean_up()
    Find_and_Select_Column ("Apply Begin Date")
        Selection.NumberFormat = "dd/mm/yyyy"
    Find_and_Select_Column ("Apply End Date")
        Selection.NumberFormat = "dd/mm/yyyy"
End Sub

Sub Theatres_DCW_clean_up()
    Range("A1").Select
    Find_and_Select_Column ("Days applied")
        Strip_Dots_Commas_Spaces_From_Selection 'warning: will destroy header
    Find_and_Select_Column ("Weeks")
        Strip_Dots_Commas_Spaces_From_Selection 'warning: will destroy header
    'Format date:
    Find_and_Select_Column ("Apply Begin Date")
        Selection.NumberFormat = "dd/mm/yyyy"
    Find_and_Select_Column ("Session Start Time")
        Selection.NumberFormat = "hh:mm"
    Find_and_Select_Column ("Session Stop Time")
        Selection.NumberFormat = "hh:mm"
Range("a1").Select
End Sub

Sub Theatres_correct_conditional_formatting()
    If Not Conditional_Format_Add("$A$5:$T$15000", "=$T5=""Addition""", "Blue", False, True, "Session Times") Then MsgBox "problem with doing format"
    If Not Conditional_Format_Add("$A$5:$T$15000", "=$T5=""Modification""", "Yellow", False, False, "Session Times") Then MsgBox "problem with doing format"
    If Not Conditional_Format_Add("$A$5:$T$15000", "=$T5=""Deletion""", "Red", True, False, "Session Times") Then MsgBox "problem with doing format"
End Sub
