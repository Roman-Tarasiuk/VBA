' In ThisWorkbook
Private Sub Workbook_NewSheet(ByVal Sh As Object)
    Application.Run "Personal.xlsb!AADefaultStyle"
End Sub

' In ThisWorkbook
Sub Auto_Open()
    Application.Run "Personal.xlsb!AADefaultStyle"
End Sub

Meta Macro:
' 'You will need to include a reference to Microsoft Visual Basic for Applications Extensibility and enable access to the VBA project module in the trust center settings.'
' https://stackoverflow.com/questions/41226485/dynamically-create-macros-specifically-udfs-from-vba-code
Sub CreateMacro()
    Dim vbComp As VBComponent
    Dim functionText As String

    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)

    functionText = "Function MyTest()" & vbCrLf
    functionText = functionText + "MsgBox " & Chr(34) & "Hello World" & Chr(34) & vbCrLf
    functionText = functionText + "End Function"

    vbComp.CodeModule.AddFromString functionText
End Sub


'
' VBScript MessageBox

x=msgbox("Your Text Here" ,0, "Your Title Here")

' 0 =OK button only
' 1 =OK and Cancel buttons
' 2 =Abort, Retry, and Ignore buttons
' 3 =Yes, No, and Cancel buttons
' 4 =Yes and No buttons
' 5 =Retry and Cancel buttons
' 16 =Critical Message icon
' 32 =Warning Query icon
' 48 = Warning Message icon
' 64 =Information Message icon
' 0 = First button is default
' 256 =Second button is default
' 512 =Third button is default
' 768 =Fourth button is default
' 0 =Application modal (the current application will not work until the user responds to the message box)
' 4096 =System modal (all applications wont work until the user responds to the message box)


' Attribute VB_Name = "Module11"
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
        Formula1:="=»—“»Õ¿"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=ÀŒ∆‹"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
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
' Attribute PasteFormat.VB_ProcData.VB_Invoke_Func = "V\n14"
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub


'
'
Sub PasteFormat2()
'
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub


'
' Unselect Cell/Area (http://www.cpearson.com/excel/UnSelect.aspx)
'
Sub UnSelectActiveCell()
    Dim R As Range
    Dim RR As Range
    For Each R In Selection.Cells
        If StrComp(R.Address, ActiveCell.Address, vbBinaryCompare) <> 0 Then
            If RR Is Nothing Then
                Set RR = R
            Else
                Set RR = Application.Union(RR, R)
            End If
        End If
    Next R
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
        .Color = 14277081
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub Color2()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10921638
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
    Dim R As Integer
    Dim S1, S2 As String
    
    R = Selection.Row
    
    Selection.EntireRow.Insert
    S1 = CStr(R) + ":" + CStr(R)
    S2 = CStr(R + 2) + ":" + CStr(R + 2)
    Rows(S2).Select
    Selection.Cut Destination:=Rows(S1)
    Range(S2).Select
    Selection.EntireRow.Delete
End Sub

Sub ZoomTo()
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
    MsgBox "Input Error. Restart the macro and enter a correct number (10-500)."
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

' The macro listens a cell change and applies the filter based on the changed value.
' Put the macro into a book's Microsoft Excel Objects sheet.
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range

    Set KeyCells = Range("V6:V6")
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then
        ActiveSheet.Range("$G$6:$V$3208").AutoFilter Field:=16, Criteria1:=Target.Value
    End If
End Sub
