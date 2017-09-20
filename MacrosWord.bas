' Maybe some code lines (like Attribute) are redundant

Attribute VB_Name = "NewMacros"
Option Explicit

Sub En()
Attribute En.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.En"
    Selection.LanguageID = wdEnglishUS
    Selection.NoProofing = False
    Application.CheckLanguage = True
End Sub

Sub Ua()
Attribute Ua.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Ua"
    Selection.LanguageID = wdUkrainian
    Selection.NoProofing = False
    Application.CheckLanguage = True
End Sub

Sub CopyFormat()
Attribute CopyFormat.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Макрос1"
' Alt+C
    Selection.CopyFormat
End Sub

Sub PasteFormat()
' Alt+V
    Selection.PasteFormat
End Sub

Sub PasteWithSelection()
' Ctrl+Alt+V
    Dim theStart As Long
    theStart = Selection.Start
    Selection.PasteAndFormat (wdFormatPlainText)
    Selection.Start = theStart
End Sub

Sub EditCopyWithoutTailSpaces()
' Ctrl+Alt+C
' Source: http://answers.microsoft.com/en-us/office/forum/office_2003-word/removing-extra-space-when-selecting-words/5748b263-d2c5-4672-a343-74ad0e22fa7f?auth=1
    Dim theStart As Long
    theStart = Selection.Start
    Dim oRng As Range
    Set oRng = Selection.Range
    While oRng.Characters.Last = Chr(32) Or oRng.Characters.Last = Chr(13)
        oRng.End = oRng.End - 1
    Wend
    oRng.Copy
    Selection.Start = theStart
End Sub

Sub PdfClipboardJoin()
Attribute PdfClipboardJoin.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.PdfClipboardJoin"
' Alt+Z
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub ForeColor1()
Attribute ForeColor1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.ForeColor"
    Selection.Font.TextColor = 15773696 ' +++
End Sub

Sub DottedUnderline()
Attribute DottedUnderline.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.DottedUnderline"
    With Selection.Font
        .underline = wdUnderlineDotted
        .underlineColor = wdColorAutomatic
    End With
End Sub

Sub DottedUnderline2()
    With Selection.Font
        .underline = wdUnderlineDotted
        .underlineColor = wdColorRed
    End With
End Sub

Sub SmallCaps()
    With Selection.Font
        .SmallCaps = True
    End With
End Sub

Sub OALDTagToFormatted(tag As String, color As Double, underline As Integer, underlineColor As Double, _
                        bold As Boolean, italic As Boolean, fontName As String)
' Used in Sub OALDCards()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<" + tag + ">"
        .Replacement.Text = "====" + tag + "===="
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "</" + tag + ">"
        .Replacement.Text = "====/" + tag + "===="
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        If color <> wdColorAutomatic Then
            .color = color
        End If
        If underline <> wdUnderlineNone Then
            .underline = underline
        End If
        If underlineColor <> 0 Then
            .underlineColor = underlineColor
        End If
        If bold <> False Then
            .bold = bold
        End If
        If italic <> False Then
            .italic = italic
        End If
        If fontName <> "" Then
            .Name = fontName
        End If
    End With
    
    With Selection.Find
        .Text = "====" + tag + "====*====/" + tag + "===="
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "====" + tag + "===="
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        With Selection.Find
        .Text = "====/" + tag + "===="
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub OALDCards()
    ' Page setup
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1)
        .RightMargin = CentimetersToPoints(1)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(29.7)
        .PageHeight = CentimetersToPoints(19.5)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
   
    ' Font and paragraph settings
    Selection.WholeStory
    Selection.Font.Size = 26
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1)
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
   
    ' Replacement 1
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " – "
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 32
        .bold = True
    End With
    With Selection.Find
        .Text = """^13""*^13"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
    ' Replacement 2
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = """^p"""
        .Replacement.Text = "^p^p^p^p^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
    ' Replacement 3
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Size = 32
        .bold = True
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 26
        .bold = False
        .italic = False
    End With
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
    ' Replacement 4
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = """"""
        .Replacement.Text = """"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
    ' Clearing begin and end of document
    Selection.HomeKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Size = 32
    Selection.Font.bold = wdToggle
    Selection.EndKey Unit:=wdStory
    Selection.TypeBackspace
   
    ' Replacement 5
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " ### "
        .Replacement.Text = "^p* "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Formatting tags
    ' OALDTagToFormatted(tag As String, color As Double, underline As Integer, underlineColor As Double, bold As Boolean, italic As Boolean, fontName As String)
    Call OALDTagToFormatted("oald8", 9792578, wdUnderlineNone, 0, False, False, "")
    Call OALDTagToFormatted("exmpl", 16750899, wdUnderlineNone, 0, True, False, "")
    Call OALDTagToFormatted("exmpla", 3329330, wdUnderlineNone, 0, False, False, "")
    Call OALDTagToFormatted("phr", wdColorAutomatic, wdUnderlineDotted, 3329330, False, False, "")
    Call OALDTagToFormatted("i", wdColorAutomatic, wdUnderlineNone, 0, False, True, "")
    Call OALDTagToFormatted("b", wdColorAutomatic, wdUnderlineNone, 0, True, False, "")
    Call OALDTagToFormatted("code", wdColorAutomatic, wdUnderlineNone, 0, False, False, "Courier New")

    ' Replacement 6
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.color = wdColorAutomatic
    With Selection.Find
        .Text = "*"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub Replace(oldValue As String, newValue As String)
'
' Used in Sub SwiftExtractClearEmail()
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = oldValue
        .Replacement.Text = newValue
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


' Sub SwiftExtractClearEmail()
' '
' ' SwiftExtractClearEmail Макрос
' '
' '
'     Call Replace("----- Переслано: Roman TARASIUK/LV/RBA-AVAL/UA дата: ^#^#.^#^#.^#^#^#^# ^#^#:^#^# -----", "")
'     Call Replace("От: from@aval.ua", "")
'     Call Replace("Кому: to@aval.ua, ", "")
'     Call Replace("Кому: to@aval.ua", "")
'     Call Replace("Дата: ^#^#.^#^#.^#^#^#^# ^#^#:^#^#", "")
'     Call Replace("Тема: Statement SWIFT format MT ^#^#^#", "")
'
'     '//
'
'     Dim length, tmp As Integer
'
'     length = Word.ActiveDocument.Characters.Count
'     Do
'         tmp = length
'         Call Replace("^p^p", "^p")
'         length = Word.ActiveDocument.Characters.Count
'     Loop While length <> tmp
'
'     '//
'
'     Call Replace("}^p", "}===^p")
'     Call Replace("^p", " *** ")
'     Call Replace("===", "^p")
'     Call Replace("^p *** ", "^p")
' End Sub


Sub SwiftExtractClearEmail()
Attribute SwiftExtractClearEmail.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.SwiftExtractClearEmail"
'
' SwiftExtractClearEmail Макрос
'
'
    MsgBox "Run this macro on a file up to 50 pages." _
        & vbNewLine & "Use 'SE.html' for further reordering of columns.", vbInformation

    Call Replace("----- Переслано: Roman TARASIUK/LV/RBA-AVAL/UA дата: ^#^#.^#^#.^#^#^#^# ^#^#:^#^# -----", "")
    Call Replace("От: from@aval.ua", "")
    Call Replace("Кому: to@aval.ua, ", "")
    Call Replace("Кому: to@aval.ua", "")
    Call Replace("Дата: ^#^#.^#^#.^#^#^#^# ^#^#:^#^#", "")
    Call Replace("Тема: Statement SWIFT format MT ^#^#^#", "")
    
    '//
    
    Dim length, tmp As Integer
    
    length = Word.ActiveDocument.Characters.Count
    Do
        tmp = length
        Call Replace("^p^p", "^p")
        length = Word.ActiveDocument.Characters.Count
    Loop While length <> tmp
    
    '//
    
    Call Replace("}^p", "}===^p")
    Call Replace("^p", "^t")
    Call Replace("===", "^p")
    Call Replace("^p^t", "^p")
End Sub


'
' Clear clipboard:
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

Sub ClearClipboard() ' Call this Sub, it uses previous arrangements
    Call FuncClearClipboard
End Sub


'
'
Sub CopyHyperlinkToClipboard()
' Add reference to %systemroot%\System32\FM20.dll
    Dim result As String
    Dim Obj As New DataObject
    
    result = Selection.Hyperlinks(1).Address
    
    Obj.SetText result
    Obj.PutInClipboard
End Sub


'
'
Sub CopyAllHyperlinksToClipboard()
    Dim hl As Hyperlink
    Dim result As String
    Dim Obj As New DataObject
    
    result = ""
    For Each hl In ActiveDocument.Hyperlinks
        result = result + hl.Address + vbCr
    Next
    
    Obj.SetText result
    Obj.PutInClipboard
End Sub


'
'
Sub ToggleBookmarks()
    ActiveWindow.View.ShowBookmarks = _
      Not ActiveWindow.View.ShowBookmarks
End Sub


'
' https://windowssecrets.com/forums/showthread.php/119477-Word-Locating-Embedded-objects-files-residing-in-the-document
Sub EmbeddedObjects()
  Dim varObj As Variant
  Dim n As Integer
  Dim names As String
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
' https://www.extendoffice.com/documents/word/750-word-select-all-objects.html
Sub EmbededObjectsSelectAll()
    Dim tempTable As Object
    Application.ScreenUpdating = False
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    
    For Each tempTable In ActiveDocument.InlineShapes
        tempTable.Range.Paragraphs(1).Range.Editors.Add wdEditorEveryone
    Next
    
    ActiveDocument.SelectAllEditableRanges wdEditorEveryone
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    Application.ScreenUpdating = True
End Sub


'
' Selection without trailing spaces
'
1. Add Class Module (named e.g. SelectionHandler):
Option Explicit

' https://msdn.microsoft.com/en-us/vba/word-vba/articles/application-windowselectionchange-event-word
' https://msdn.microsoft.com/VBA/Word-VBA/articles/using-events-with-the-application-object-word
Public WithEvents appWord As Word.Application

' http://excelrevisited.blogspot.com/2012/06/endswith.html
Public Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     Dim ss As String
     endingLen = Len(ending)
     EndsWith = (Right(UCase(str), endingLen) = UCase(ending))
End Function

Private Sub appWord_WindowSelectionChange(ByVal Sel As Selection)
 Dim diff As Integer
 If EndsWith(Sel.Text, " ") And (Len(Sel.Text) > 1) Then
    diff = Len(Sel.Text) - Len(Trim(Sel.Text))
    If diff <> Len(Sel.Text) Then
        Sel.End = Sel.End - diff
    End If
 End If
End Sub
2. Add Normal Module:
Option Explicit

Dim X As New SelectionHandler

Sub AAHandleSelection()
    Set X.appWord = Word.Application
End Sub
