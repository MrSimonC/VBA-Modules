Attribute VB_Name = "Functions"
Function CheckFileIsOpen(chkSumfile As String) As Boolean
'Check if a file is open
'http://www.mrexcel.com/forum/excel-questions/431458-visual-basic-applications-code-open-excel-file-only-if-not-already-open.html
'Input: FileName.ext
    On Error Resume Next
    CheckFileIsOpen = (Workbooks(chkSumfile).Name = chkSumfile)
    On Error GoTo 0
End Function

Function SheetExists(strWSName As String) As Boolean
'Check if tab name exists in current workbook
'Input: Sheet/Tab name
'Boolean function assumed to be False unless set to True
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(strWSName)
    If Not ws Is Nothing Then SheetExists = True
End Function

Public Function FileFolderExists(strFullPath As String) As Boolean
'Check if a file or folder exists
'www.excelguru.ca
'Input: c:\full\path.ext
    On Error Resume Next
    If strFullPath = "" Then 'sac addition
        FileFolderExists = False
        Exit Function
    End If
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    On Error GoTo 0
End Function
 
Function ChooseFile(dialogueTitle As String, PathToStart As String, ByRef FilePath As String) As Boolean
'Show file chooser dialogue box
'Returns: Path of selected file or "" if nothing selected
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim FileChosen As Integer 'get the number of the button chosen
    
    fd.Title = dialogueTitle
    fd.InitialFileName = PathToStart     'initially loaded folder
    fd.InitialView = msoFileDialogViewDetails   'the folder view
    fd.Filters.Clear
    fd.Filters.Add "All Files", "*.*"
    fd.Filters.Add "CSV Files", "*.csv"
    fd.Filters.Add "Excel Workbooks", "*.xlsx"
    fd.Filters.Add "Excel Macros", "*.xlsm"
    fd.FilterIndex = 1  'default filter (above) to show
    'fd.ButtonName = "OK button"    'text appearing on the OK Button
    FileChosen = fd.Show

    If FileChosen <> -1 Then 'didn't choose anything (clicked on CANCEL)
        FilePath = ""
        ChooseFile = False
    Else
        FilePath = fd.SelectedItems(1) 'display name and path of file chosen
        ChooseFile = True
    End If
End Function

Function GetFileName(FullPath As String, ByRef FileName As String) As Boolean
'Gets just the filename from a full path
'Returns: Filename.ext
    Dim StrFind As String
    
    If FullPath = "" Then
        GetFileName = False
        Exit Function
    End If
    
    Do Until Left(StrFind, 1) = "\"
        iCount = iCount + 1
        StrFind = Right(FullPath, iCount)
        If iCount = Len(FullPath) Then
            GetFileName = False
            Exit Function
        End If
    Loop

    FileName = Right(StrFind, Len(StrFind) - 1)
    GetFileName = True
End Function

Function Conditional_Format_Add(Range_to_apply As String, Formula As String, Colour As String, Optional Strikethrough As Boolean = False, Optional Wipe_all_formatting As Boolean = False, Optional SheetName As String = "", Optional ChangeFont As Boolean = True, Optional FontColour As String = "Black")
'Conditional Formatting
    Dim CurrentCell As String
    Dim CurrentSheet As String
    
    'Record where we are
    CurrentSheet = ActiveSheet.Name
    CurrentCell = ActiveCell.Address
    
    'Switch to the named sheet, or just use current one
    If SheetName <> "" Then Sheets(SheetName).Select
    
    'kill all current formatting
    If Wipe_all_formatting Then Cells.FormatConditions.Delete
    
    'Do the rule addition
    Range(Range_to_apply).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:=Formula
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    'Cell colour
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        If IsNumeric(Colour) Then
            .Color = Colour
        Else
            Select Case Colour
                Case "White"
                    .Color = 16777215
                Case "Red"
                    .Color = 255
                Case "Blue"
                    .Color = 15773696
                Case "Yellow"
                    .Color = 65535
                Case "Green"
                    .Color = RGB(51, 204, 51)
            Case Else
                Conditional_Format_Add = False
                Exit Function
            End Select
        End If
        .TintAndShade = 0
    End With
    'Font Colour
    If ChangeFont Then
        With Selection.FormatConditions(1).Font
            If IsNumeric(FontColour) Then
                .Color = FontColour
            Else
                Select Case FontColour
                    Case "Black"
                        .Color = 1
                    Case "Red"
                        .Color = 255
                    Case "Blue"
                        .Color = 15773696
                    Case "Yellow"
                        .Color = 65535
                    Case "Green"
                        .Color = RGB(51, 204, 51)
                Case Else
                    Conditional_Format_Add = False
                    Exit Function
                End Select
            End If
        End With
    End If
    Selection.FormatConditions(1).Font.Strikethrough = Strikethrough
    Selection.FormatConditions(1).StopIfTrue = False
    
    Sheets(CurrentSheet).Select
    Range(CurrentCell).Select
    Conditional_Format_Add = True
End Function

Sub FillColor()
    MsgBox ActiveCell.Interior.Color
End Sub

Sub FontColour()
    MsgBox ActiveCell.Font.ColorIndex
End Sub

Sub Conditional_Format_Test()
    Conditional_Format_Add "A1:E8", "=$E1=2", "Yellow", True
End Sub

Function FindCellRef(rangeToSearch As Range, item As String) As String
'Find a cell's range in a row/column
'Input: Range to search [e.g. Range("4:4")], the item to look for.
'Returns: Cell reference
        rangeToSearch.Select  'Find "Template Build Name" in entire sheet, then Ctrl+Shift down and left
        Cells.Find(item, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Select
        FindCellRef = ActiveCell.Address
End Function

Function BottomLeftUsedRange() As String
'Find a cell's bottom left last used cell reference
'Returns: Cell reference
   
    ActiveCell.SpecialCells(xlLastCell).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    BottomLeftUsedRange = ActiveCell.Address
End Function


Function Levenshtein(ByVal string1 As String, ByVal string2 As String) As Long
'Levenshtein Distance - returns long number, lower = closest match, higher or 0 = not a good match
'http://stackoverflow.com/questions/4243036/levenshtein-distance-in-excel

Dim i As Long, J As Long
Dim string1_length As Long
Dim string2_length As Long
Dim distance() As Long

'simon addition: matches closer if you push both to lowercase
string1 = LCase(string1)
string2 = LCase(string2)

string1_length = Len(string1)
string2_length = Len(string2)
ReDim distance(string1_length, string2_length)

For i = 0 To string1_length
    distance(i, 0) = i
Next

For J = 0 To string2_length
    distance(0, J) = J
Next

For i = 1 To string1_length
    For J = 1 To string2_length
        If Asc(Mid$(string1, i, 1)) = Asc(Mid$(string2, J, 1)) Then
            distance(i, J) = distance(i - 1, J - 1)
        Else
            distance(i, J) = Application.WorksheetFunction.Min _
            (distance(i - 1, J) + 1, _
             distance(i, J - 1) + 1, _
             distance(i - 1, J - 1) + 1)
        End If
    Next
Next

Levenshtein = distance(string1_length, string2_length)
End Function

Function Levenshtein_return_best_match(toMatch As String, ByRef Possibilities() As String) As String
'Return the best match from toMatch against the Possibilities array
    Dim counter As Integer
    Dim bestMatch As String
    Dim lowestNumber As Integer
    counter = 0
    bestMatch = ""
    lowestNumber = 10000
    
    ReDim Preserve Possibilities(UBound(Possibilities), 2) 'resize array appropriately
    
    'Populate Levenshtein scores & Work out best possibility
    For counter = 1 To UBound(Possibilities)
        Possibilities(counter, 2) = Levenshtein(toMatch, Possibilities(counter, 1)) 'value of match (0=none, lowest=best)
        If Possibilities(counter, 2) <> 0 And Possibilities(counter, 2) < lowestNumber Then
            lowestNumber = Possibilities(counter, 2)
            bestMatch = Possibilities(counter, 1)
        End If
    Next
    
    Levenshtein_return_best_match = bestMatch

End Function

Function Put_Selection_into_Array(ByRef arrayOfItems() As String)
    Dim counter As Integer: counter = 0
    
    ReDim arrayOfItems(Selection.Count)
    For Each itemToCheck In Selection
            counter = counter + 1
            arrayOfItems(counter) = itemToCheck.Value
        Next itemToCheck
        
    Put_Selection_into_Array = arrayOfItems()
End Function

Function Strikethrough(r As Range)
    'sees if cell has strikethrough or not
    'use with =PERSONAL.XLSB!Strikethrough(A1)
    Strikethrough = r.Font.Strikethrough
End Function

Function ConcatRange(inputRange As Range, Optional delimiter As String) As String
'taken from online, concats an entire column into a cell
    Dim oneCell As Range
    With inputRange
        If Not (Application.Intersect(.Parent.UsedRange, .Cells) Is Nothing) Then
            For Each oneCell In Application.Intersect(.Parent.UsedRange, .Cells)
                If oneCell.Text <> vbNullString Then
                    ConcatRange = ConcatRange & delimiter & oneCell.Text
                End If
            Next oneCell
            ConcatRange = Mid(ConcatRange, Len(delimiter) + 1)
        End If
    End With
End Function

Function GetAddress(HyperlinkCell As Range)
    GetAddress = Replace(HyperlinkCell.Hyperlinks(1).Address, "mailto:", "")
End Function

Public Function GetURL(c As Range) As String
    On Error Resume Next
    GetURL = c.Hyperlinks(1).Address
End Function

Function simpleCellRegex(Myrange As Range) As String
    'http://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String

    'sac - continue trying this out
    'original example:
    'strPattern = "^[0-9]{1,3}"
    'sac testing:
    'strPattern = "^[a-zA-Z0-9\'\-\.\ ]*$"
    strPattern = "^[a-zA-Z]*"

    If strPattern <> "" Then
        strInput = Myrange.Value
        strReplace = ""

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.test(strInput) Then
            simpleCellRegex = regEx.Replace(strInput, strReplace)
        Else
            simpleCellRegex = "Not matched"
        End If
    End If
End Function

Function HyperlinkValid(ByVal strUrl As String) As Boolean
    'http://stackoverflow.com/questions/1118221/sort-dead-hyperlinks-in-excel-with-vba
    
    'Requires: Tools, References, Microsoft XML V3 (or above)
    Dim oHttp As New MSXML2.XMLHTTP30

    On Error GoTo ErrorHandler
    oHttp.Open "HEAD", strUrl, False
    oHttp.send

    If Not oHttp.Status = 200 Then HyperlinkValid = False Else HyperlinkValid = True

    Exit Function

ErrorHandler:
    CheckHyperlink = False
End Function

Function FileExists(cell As Range)
    If Dir(cell.Value) <> "" Then FileExists = True Else FileExists = False
End Function


