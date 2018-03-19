Attribute VB_Name = "Tools"
Sub Trim_Selected_Cells()
    Dim cell As Excel.Range
    
    On Error GoTo ErrH:
    'disable other events - like "on change" - sac- very useful for fighting Cerner DCWs! ;)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    
    For Each cell In Selection
        If cell.HasFormula = False Then
            cell = Trim(cell)
        End If
        Application.StatusBar = "Processing: " & cell.Row & " of " & Selection.Count
    Next cell
    
    'turn on events again
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub
    
ErrH:
    'turn on events again
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Problem occured - please investigate code! ... Sorry!", vbCritical
End Sub

Sub Reset_Used_Range()
    'http://www.mvps.org/dmcritchie/excel/lastcell.htm
    Dim sh As Worksheet, x As Long
    For Each sh In ActiveWorkbook.Worksheets
    x = sh.UsedRange.Rows.Count
    Next sh
End Sub

Sub Delete_Unused_Cells_And_Reset_Used_Range_All_Sheets()
    'from http://www.contextures.com/xlfaqApp.html#Unused
    'Note: This code may not work correctly if the worksheet contains merged cells. To check your worksheet, you can run the TestForMergedCells code.
    Dim myLastRow As Long
    Dim myLastCol As Long
    Dim wks As Worksheet
    Dim dummyRng As Range
    
    For Each wks In ActiveWorkbook.Worksheets
      With wks
        myLastRow = 0
        myLastCol = 0
        Set dummyRng = .UsedRange
        On Error Resume Next
        myLastRow = _
          .Cells.Find("*", After:=.Cells(1), _
            LookIn:=xlFormulas, LookAt:=xlWhole, _
            SearchDirection:=xlPrevious, _
            SearchOrder:=xlByRows).Row
        myLastCol = _
          .Cells.Find("*", After:=.Cells(1), _
            LookIn:=xlFormulas, LookAt:=xlWhole, _
            SearchDirection:=xlPrevious, _
            SearchOrder:=xlByColumns).Column
        On Error GoTo 0
    
        If myLastRow * myLastCol = 0 Then
            .Columns.Delete
        Else
            .Range(.Cells(myLastRow + 1, 1), _
              .Cells(.Rows.Count, 1)).EntireRow.Delete
            .Range(.Cells(1, myLastCol + 1), _
              .Cells(1, .Columns.Count)).EntireColumn.Delete
        End If
      End With
    Next wks
End Sub

Sub Find_and_Select_Column(SearchTerm As String)    'searches in currently active row
    ActiveCell.Rows.EntireRow.Select
    Cells.Find(SearchTerm, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.EntireColumn.Select
End Sub

Sub Find_and_Select_Range(SearchTerm As String, ColumnToCheckForEndOfData As String) 'searches active row (for top right marker) then bottom of sheet up in column (as bottom left marker)
    ActiveCell.Rows.EntireRow.Select
    Cells.Find(SearchTerm, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Select
    TopRight = ActiveCell.Address

    BottomOfRowANumber = Range(ColumnToCheckForEndOfData & Rows.Count).End(xlUp).Row

    Range(TopRight, ColumnToCheckForEndOfData & BottomOfRowANumber).Select
End Sub

Sub Strip_Dots_Commas_Spaces_From_Selection()   'Warning: if you have the entire column selected, this will affect the header
    Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Sub moveCurrentColumnToRight()  'assumes you're on a column you want to move
    ActiveCell.EntireColumn.Select
    Selection.Cut
    Range("A1").Select
    Selection.End(xlToRight).Select
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.Insert Shift:=xlToRight   'does the "insert cut cells right click command
    Range("A1").Select
End Sub

Sub Bank_statement_create_criteria()
    Dim strRangeStartAddress As String  'temporary string to hold the "starting range"
    Dim StatementCell As Excel.Range
    Dim CategoriesRange As Variant 'create array of all lookup table
    Dim StatementRangeToSearchThrough As Range 'Bank statement range to search through
    Dim CurrentSheet As String
    Dim CurrentCell As String
    
    CurrentSheet = ActiveSheet.Name
    CurrentCell = ActiveCell.Address
    
    'Find list of Items on the bank statement to search through
    Sheets("Statement").Select
    strRangeStartAddress = Rows("1:1").Find(What:="Description", After:=Cells(1, 1), _
        LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, MatchCase:=True)(2, 1).Address
    'Set range to search through by doing down: assumes no empty cells
    Set StatementRangeToSearchThrough = Range(strRangeStartAddress, Range(strRangeStartAddress).End(xlDown))
    
    'Look for "Lookup table" cell reference. (2, 1)= move down one
    Sheets("List of Shops").Select
    strRangeStartAddress = Rows("1:1").Find(What:="Shop Name", After:=Cells(1, 1), _
        LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, MatchCase:=True)(2, 1).Address
    'Set category range by doing down and right
    CategoriesRange = Range(strRangeStartAddress, Range(strRangeStartAddress).End(xlDown).End(xlToRight))
    
    'For each statement entry
    Sheets("Statement").Select
    For Each StatementCell In StatementRangeToSearchThrough
        'and for each column 1 (shop name) in lookup table/array
        For i = 1 To UBound(CategoriesRange)
            If InStr(1, StatementCell, CategoriesRange(i, 1), vbTextCompare) <> 0 Then
                Cells(StatementCell.Row, 7).Value = CategoriesRange(i, 2)
                'ActiveCell.Offset(StatementCell.Row, 3).Value = CategoriesRange(i, 2)
            End If
        Next
    Next StatementCell
    
    'Sort out auto-making unique list of Categories from List of Shops
    'Clear out Summary sheet current data
    Sheets("Summary").Select
        Cells.Find(What:="Category", After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        Range(Selection(2, 1), Selection.End(xlDown)).Select
        Selection.ClearContents
        'Go get unique categories including 'Category' title
    Sheets("List of Shops").Select
        Columns("B:B").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("C1" _
            ), Unique:=True
        Range("C1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        'paste it in Summary sheet
    Sheets("Summary").Select
        Cells.Find(What:="Category", After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        'deselect things
    Sheets("List of Shops").Select
        Columns("C:C").Select
        Selection.ClearContents
        Range("B1").Select
    Sheets("Summary").Select
        Range("A8").Select
        
    Sheets(CurrentSheet).Select
    Range(CurrentCell).Select
    
End Sub

Sub Find_Cosest_Resource_with_Levenshtein() 'Takes a list of strings and in the right column types the closest match
    Dim freeText As String
    Dim resourceNames() As String 'make array of possible matches
    Dim counter As Integer
    Dim ResourcesFilePath As String
    Dim ResourcesFileName As String
    Dim ResourcesCellToFind As String
    Dim listToCheck() As String
    
    Erase resourceNames()
    ResourcesFilePath = "C:\Cerner Audit\SPFIT_RESOURCE_EXTRACT.csv"
    ResourcesCellToFind = "Resource"
    
    'Get list if many items are selected
    Put_Selection_into_Array listToCheck()
    
    'Check the resources extract is there
    If Not FileFolderExists(ResourcesFilePath) Then
        MsgBox "Can't see resources file at " & ResourcesFilePath
        Exit Sub
    End If
    
    'Check we've something to work with!
    If ActiveCell.Value = "" Then
        MsgBox "You're running the macro on an empty cell - please click in a cell with content in before running."
        Exit Sub
    End If
    
    For i = 1 To UBound(listToCheck())
        counter = 0
        'Set first freetext value
        freeText = listToCheck(i)
        
        'get list of resources
        Application.ScreenUpdating = False
            Workbooks.Open ResourcesFilePath, ReadOnly:=True
            Find_and_Select_Range ResourcesCellToFind, "A"
            
            'put possibilities text into array
            ReDim resourceNames(Selection.Count, 2) 'resize array appropriately
            For Each cell In Selection
                counter = counter + 1
                resourceNames(counter, 1) = cell.Value    'name of potential match
            Next cell
            
            'close resources file
            GetFileName ResourcesFilePath, ResourcesFileName
            Workbooks(ResourcesFileName).Close SaveChanges:=False
        Application.ScreenUpdating = True
        
        ActiveCell.Offset(0, 1).Select  'move right
        ActiveCell.Value = Levenshtein_return_best_match(freeText, resourceNames())
        ActiveCell.Offset(1, 0).Select 'move down
        ActiveCell.Offset(0, -1).Select  'move left
    Next
End Sub

Sub Conditional_Format_AddModifyDelete(Range As String, ModifiedColumn As String)
    Conditional_Format_Add Range, "=$" & ModifiedColumn & "1=""Addition""", "Blue", False, True
    Conditional_Format_Add Range, "=$" & ModifiedColumn & "1=""Modification""", "Yellow", False, False
    Conditional_Format_Add Range, "=$" & ModifiedColumn & "1=""Deletion""", "Red", True, False
End Sub

Sub Conditional_Format_AddModifyDelete_with_Values()
    Conditional_Format_AddModifyDelete "A:G", "G"
End Sub

Sub Conditional_Format_2del()
    Conditional_Format_Add "A:T", "=$T1=1", "Blue", False, True
    Conditional_Format_Add "A:T", "=$T1=2", "Red", False
    Conditional_Format_Add "A:T", "=$T1=3", "Green", False
    Conditional_Format_Add "A:T", "=$T1=4", "Yellow", False
    Conditional_Format_Add "A:T", "=$T1=5", "Orange", False
    Conditional_Format_Add "A:T", "=$T1=6", "Brown", False
End Sub

Sub Conditional_Format_RAG()
    Dim applyToRange As String
    applyToRange = "$A:$A"
    Conditional_Format_Add applyToRange, "=A1=""R""", "Red", False, True, , True, "Red"
    Conditional_Format_Add applyToRange, "=A1=""A""", "Yellow", False, False, , True, "Yellow"
    Conditional_Format_Add applyToRange, "=A1=""G""", "Green", False, False, , True, "Green"
    'Conditional_Format_Add applyToRange, "=F1=""B""", "Black", False, False, , True, "Yellow"
    Conditional_Format_Add applyToRange, "=A1=""N""", "Red", False, False, , True, "Red"
    Conditional_Format_Add applyToRange, "=A1=""Y""", "Green", False, False, , True, "Green"
End Sub

Sub Conditional_Format_CurrentPrevious_for_Comparisons()
    Dim applyToRange As String
    applyToRange = "$A:$H"
    Conditional_Format_Add applyToRange, "=$C1=""Previous""", "White", False, True, "", True, "Red"
    Conditional_Format_Add applyToRange, "=$C1=""Current""", "White", False, False, "", True, "Blue"
End Sub


Sub Conditional_Format_end_slot_for_cerner_to_lorenzo()
    'sortLorColumns
    Freeze_Top_Row
    Dim applyToRange As String
    previous_column = "AP"
    end_slot_column = "AQ"
    applyToRange = "$" + end_slot_column + ":$" + end_slot_column
    'Next slot requires an inactive slot:
    Conditional_Format_Add applyToRange, "=and($a1=$a2,$c1=$c2,$d1=$d2,$" + end_slot_column + "1<$" + previous_column + "2, $" + end_slot_column + "1<>$" + previous_column + "2)", "Red", False, True, "", False
    'Column separators
    Conditional_Format_Add "$A:$DD", "=A1=""---""", "Yellow", False, False, "", False
    'Session Separators (AE=Session Name, AG=From Date, 10092543=soft yellow colour)
    Conditional_Format_Add "$AD$2:$DD$50000", "=OR($AE2<>$AE1,and($AE2=$AE1,$AG2<>$AG1))", 10092543, False, False, "", False
    'Double spaces
    Conditional_Format_Add "$A:$ZZ", "=NOT(ISERR(FIND(""  "",A1)))", "White", False, False, "", True, "Red"
    'Grouping
    'UnGroup ("a:www")
    Group ("a:ad")
End Sub

Sub Conditional_Domain_or_DCS()
    Conditional_Format_Add "A:N", "=$A1=""Domain""", "White", , True, "", True, "Red"
    Conditional_Format_Add "A:N", "=$A1=""DCS""", "White", , , "", True, "Green"
End Sub

Sub CreateSheetsFromAList()
    Dim MyCell As Range, MyRange As Range
    
    Set MyRange = Sheets("Sheet1").Range("A1")
    Set MyRange = Range(MyRange, MyRange.End(xlDown))

    For Each MyCell In MyRange
        Sheets.Add After:=Sheets(Sheets.Count) 'creates a new worksheet
        Sheets(Sheets.Count).Name = Left(MyCell.Value, 30) ' renames the new worksheet
    Next MyCell
End Sub

Sub Freeze_Top_Row()
ActiveWindow.FreezePanes = False
Range("A2").Select
ActiveWindow.FreezePanes = True
End Sub

Sub Group(Columns)
Range(Columns).Select
Selection.Columns.Group
End Sub

Sub UnGroup(Columns)
Range(Columns).Select
Selection.Columns.UnGroup
End Sub

Sub replace_null()
Application.EnableEvents = False
Cells.Replace What:="NULL", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
Application.EnableEvents = True
MsgBox "All Nulls Replaced."
End Sub

Sub Combine_data_from_all_sheets_into_one()
    'http://excel.tips.net/T003005_Condensing_Multiple_Worksheets_Into_One.html
    Dim J As Integer

    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add ' add a sheet in first place
    Sheets(1).Name = "Combined"

    ' copy headings
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    ' work through sheets
    For J = 2 To Sheets.Count ' from sheet 2 to last sheet
        Sheets(J).Activate ' make the sheet active
        Range("A1").Select
        'Selection.CurrentRegion.Select ' select all cells in this sheets
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select    'sac: select whole sheet

        ' select all lines except title
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select

        ' copy cells selected in the new sheet on last line
        Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
    Next
End Sub


Sub TraverseFolderInOutlook()
' https://msdn.microsoft.com/en-us/library/office/ff870566(v=office.14).aspx
Dim outlook As Object
Set outlook = CreateObject("Outlook.Application")
Set mapi = outlook.GetNamespace("MAPI")
Set MAPIFolder = mapi.GetDefaultFolder(6)
Set Folders = MAPIFolder.Folders

Dim folder_to_traverse As Object
For Each flr In Folders
    If flr.Name = "Test" Then
        Set folder_to_traverse = flr
        MsgBox folder_to_traverse.Name
    End If
Next

Dim item As Object
For Each item In folder_to_traverse.Items
    If InStr(item.Subject, "Undeliverable") Then
        MsgBox item.Body
    End If
Next
End Sub


Sub ddmmyyyy_hhmm()
    Selection.NumberFormat = "m/d/yyyy h:mm"
End Sub


Sub ColorCellsByHex()
  ' Changes background of cell to the #Hex colour value
  Dim rSelection As Range, rCell As Range, tHex As String

  If TypeName(Selection) = "Range" Then
  Set rSelection = Selection
    For Each rCell In rSelection
      tHex = Mid(rCell.Text, 6, 2) & Mid(rCell.Text, 4, 2) & Mid(rCell.Text, 2, 2)
      rCell.Interior.Color = WorksheetFunction.Hex2Dec(tHex)
    Next
  End If
End Sub

