>>>>
'Private Sub Workbook_Open()
'End Sub


'
' ClearClipboard() with helper functions
'
Option Explicit
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

Private Function FuncClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Function
 
Sub ClearClipboard()
    Call FuncClearClipboard
End Sub


'
'
Sub TrueFalseConditionalFormatting()
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=TRUE"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -11480942
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=FALSE"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -8081164
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=≈Œÿ»¡ ¿(RC)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 12566463
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub


'
'
Function Last(str As String, strMatch As String)
    Last = InStrRev(str, strMatch)
End Function


'
'
Sub PasteFormat()
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub


'
' Unselect Cell/Area (http://www.cpearson.com/excel/UnSelect.aspx)
'
Sub UnSelectActiveCell()
    Dim r As Range
    Dim RR As Range
    For Each r In Selection.Cells
        If StrComp(r.Address, ActiveCell.Address, vbBinaryCompare) <> 0 Then
            If RR Is Nothing Then
                Set RR = r
            Else
                Set RR = Application.Union(RR, r)
            End If
        End If
    Next r
    If Not RR Is Nothing Then
        RR.Select
    End If
End Sub

Sub UnSelectCurrentArea()
    Dim Area As Range
    Dim RR As Range
    
    For Each Area In Selection.Areas
        If Application.Intersect(Area, ActiveCell) Is Nothing Then
            If RR Is Nothing Then
                Set RR = Area
            Else
                Set RR = Application.Union(RR, Area)
            End If
        End If
    Next Area
    If Not RR Is Nothing Then
        RR.Select
    End If
End Sub


'
'
Sub CellTextFormat()
    Selection.NumberFormat = "@"
End Sub

Sub CellNumberFormat()
    Selection.NumberFormat = "0"
End Sub

Sub CellNumber2DecPlacFormat()
    Selection.NumberFormat = "0.00"
End Sub

Sub CellDateFormat()
    Selection.NumberFormat = "dd/mm/yyyy"
End Sub

Sub CellTimeFormat()
    Selection.NumberFormat = "hh:mm:ss"
End Sub

Sub CellDateTimeFormat()
    Selection.NumberFormat = "dd/mm/yyyy hh:mm:ss"
End Sub

Sub CellCurrencyFormat()
    Selection.NumberFormat = "#,##0.00_?"
End Sub


'
'
Sub Color1()
' Color format: BGR -> decimal
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub Color2()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 14277081
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub ColorClear()
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


'
'
Sub UniqueValues()
    Selection.AdvancedFilter Action:=xlFilterInPlace, Unique:=True
    ActiveWindow.SmallScroll Down:=-3
End Sub


'
'
Sub CopyHyperlinkToClipboard()
' Add reference to %systemroot%\System32\FM20.dll
    Dim a As String
    Dim Obj As New DataObject
    
    a = Selection.Hyperlinks(1).Address
    
    Obj.SetText a
    Obj.PutInClipboard
End Sub


'
'
Sub EmbeddedObjects()
  Dim varObj As Variant
  Dim n As Integer
  Dim names As String
  Dim Q As DataObject
  Q = ActiveSheet
  For Each varObj In ActiveDocument.InlineShapes
    If varObj.Type = wdInlineShapeEmbeddedOLEObject Then
      n = n + 1
      names = names + CStr(n) + ") " + varObj.OLEFormat.IconLabel + vbCrLf
    End If
  Next varObj
  For Each varObj In ActiveDocument.Shapes
    If varObj.Type = msoEmbeddedOLEObject Then
      n = n + 1
    End If
  Next varObj
  MsgBox names + "--" + vbCrLf + "Total: " + CStr(n) + " files and shapes", vbInformation
End Sub


'
'
Sub TestForUnsavedChanges()
    If ActiveWorkbook.Saved = False Then
        MsgBox "This workbook contains unsaved changes."
    Else
        MsgBox "The workbook is saved."
    End If
End Sub


Sub SaveWithCheck()
    If ActiveWorkbook.Saved = False Then
        ActiveWorkbook.Save
    End If
End Sub

Sub SaveAll()
' https://www.extendoffice.com/documents/excel/2971-excel-save-all-open-files.html
    Dim xWb As Workbook
    For Each xWb In Application.Workbooks
        If Not xWb.ReadOnly _
                And Windows(xWb.Name).Visible _
                And xWb.Saved = False _
        Then
            xWb.Save
        End If
    Next
End Sub


Sub SecurityDDE()
    Application.IgnoreRemoteRequests = False
End Sub


Sub AADefaultStyle()
    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.ColumnWidth = 6.33
    Selection.RowHeight = 15
    Range("A1").Select
End Sub


Sub Swap2Rows()
    Dim r As Integer
    Dim S1, S2 As String
    
    r = Selection.Row
    
    Selection.EntireRow.Insert
    S1 = CStr(r) + ":" + CStr(r)
    S2 = CStr(r + 2) + ":" + CStr(r + 2)
    Rows(S2).Select
    Selection.Cut Destination:=Rows(S1)
    Range(S2).Select
    Selection.EntireRow.Delete
End Sub


Sub ZoomTo()
'
' ZoomTo Macro
'

'
    On Error GoTo TheError
    Dim PercentageStr As String
    Dim PercentageNum As Integer
    
    PercentageStr = InputBox("Enter zoom percentage:", , ActiveWindow.Zoom)
    
    If PercentageStr = vbNullString Then
        Exit Sub
    End If
    
    PercentageNum = PercentageStr
    
    If PercentageNum < 10 Then
        PercentageNum = 10
    ElseIf PercentageNum > 400 Then
        PercentageNum = 400
    End If
    
    ActiveWindow.Zoom = PercentageNum
    GoTo TheEnd
    
TheError:
    MsgBox "Input Error. Try again and enter a correct number (10-500)."
    Exit Sub
TheEnd:
End Sub


Sub FormulaStyle()
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub


Sub NumberSeparator()
    Application.UseSystemSeparators = Not Application.UseSystemSeparators
End Sub


Sub ExplorePath()
    Shell Environ("windir") & "\Explorer.exe " & ActiveWorkbook.Path, vbMaximizedFocus
End Sub


Sub DeleteTextBoxes()
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
End Sub


Sub SimpleSeries()
    Selection.DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, _
        Step:=1, Trend:=False
End Sub


Sub SimpleSeriesRight()
    Selection.DataSeries Rowcol:=xlRows, Type:=xlLinear, Date:=xlDay, _
        Step:=1, Trend:=False
End Sub


Sub BuildField1()
    Dim S, F As String
    Dim StartPos, EndPos As Integer
    
    S = Cells(7, 11).FormulaR1C1
    StartPos = InStr(4, S, "[")
    EndPos = InStr(7, S, "]")
    
    F = "=Exported!R[" _
        + CStr(CInt(Mid(S, StartPos + 1, EndPos - StartPos - 1)) - 1) _
        + "]C"
    
    Cells(8, 11).FormulaR1C1 = F
End Sub

Sub BuildField2()
    Dim F As String
    
    F = "=R[-5]C=Exported!R[" _
        + CStr(CInt(Cells(10, 2).Value) - 7) _
        + "]C"
    
    Cells(7, 1).FormulaR1C1 = F
End Sub

Sub GoHome()
    SendKeys ("^{HOME}")
End Sub

Sub BuildField1ver2()
    Dim S, ResultFormula, r As String
    Dim i, length As Integer
    Dim cols As Integer
    
    cols = 18
    
    r = Cells(9, 2).Value
    
    For i = 1 To cols
        S = Cells(4, i).FormulaR1C1
        
        ' 1. Compare values.
        '
        'ResultFormula = Replace(S, "[-3]", "[-4]", 1, 1)
        'ResultFormula = Replace(ResultFormula, "[-3]", R, 1, 1)
        'Cells(6, i).FormulaR1C1 = ResultFormula
        
        ' 2. Copy values.
        '
        length = Len(S)
        If length > 8 Then
            ResultFormula = Mid(S, 8, length - 7)
            ResultFormula = Replace(ResultFormula, "[-3]", r, 1, 1)
            Cells(7, i).FormulaR1C1 = ResultFormula
        End If
    Next i
End Sub

Function AllAreTrue(rng As Range) As Boolean
    Dim cell As Range
    
    AllAreTrue = True
    For Each cell In rng
        If cell.Value = False Then
            AllAreTrue = False
            Exit For
        End If
    Next cell
End Function

>>>>
Option Explicit

Sub CopyCellValueToClipboard()
'
' CopyCellValueToClipboard Macro
'
' Keyboard Shortcut: Ctrl+q
'
    Dim MyData As Object
    Set MyData = New DataObject
    MyData.SetText Selection.Text
    MyData.PutInClipboard
End Sub


Sub CopyCellValueToClipboard2()
    ' Ctrl + ?
    Call CopyCellValueToClipboard
End Sub

>>>>
Option Explicit

Sub FillUp()
'
' FillUp Macro
'
' Keyboard Shortcut: Ctrl+Shift+D
'
    Selection.FillUp
End Sub

Sub FillLeft()
'
' FillLeft Macro
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    Selection.FillLeft
End Sub

>>>>
Option Explicit

Sub PasteValue()
'
' PasteValue Macro
'
' Keyboard Shortcut: Ctrl+e
'
    Selection.NumberFormat = "#,##0.00_?"
    ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, _
        DisplayAsIcon:=False
End Sub




' Filtering based on a cell's value.
' Uses a Table instead of a range.
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range

    Set KeyCells = Range("D2:D2")
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
        Dim strFilter As String
        strFilter = "*" & [D2] & "*"
        Debug.Print strFilter
        ActiveSheet.ListObjects("TableBins").Range.AutoFilter _
            Field:=1, _
            Criteria1:=strFilter, _
            Operator:=xlFilterValues
    End If
End Sub


Sub ListSheets()
    Dim ws As Worksheet
    Dim rStart, r, c As Integer

    rStart = ActiveCell.Row
    r = ActiveCell.Row
    c = ActiveCell.Column

    For Each ws In Worksheets
        If Not ws.Name = ActiveSheet.Name Then
            ActiveSheet.Cells(r, c) = ws.Name
            ' Adding hyperlink to each sheet.
            'ActiveSheet.Cells(r, c).Select
            'ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
            '    ws.Name & "!R1C1", TextToDisplay:=ws.Name
            r = r + 1
        End If
    Next ws
    
    ' Selecting last cell or created list.
    'Cells(r - 1, c).Select
    'Range(Cells(rStart, c), Cells(r - 1, c)).Select
End Sub


' Can be used with the previous macro.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Intersect(Range("C3:C12"), Target) Is Nothing Then
        Sheets(ActiveCell.Text).Select
    End If
End Sub


Sub Combine()
' https://www.extendoffice.com/documents/excel/1184-excel-merge-multiple-worksheets-into-one.html
    Dim J As Integer
    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add
    Sheets(1).Name = "Combined"
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")
    
    For J = 2 To Sheets.Count
        Sheets(J).Activate
        Range("A1").Select
        Selection.CurrentRegion.Select
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
        Selection.Copy Destination:=Sheets(1).Range("A1048576").End(xlUp)(2) ' Use "A65536" for old xls.
    Next
End Sub


Function FindR(c As String, m As String)
    FindR = InStrRev(c, m)
End Function
