>>>>
Option Explicit


Sub En()
    Selection.LanguageID = wdEnglishUS
    ' Selection.NoProofing = False
    Application.CheckLanguage = True
End Sub


Sub Ua()
    Selection.LanguageID = wdUkrainian
    ' Selection.NoProofing = False
    Application.CheckLanguage = True
End Sub


Sub Ru()
    Selection.LanguageID = wdRussian
    ' Selection.NoProofing = False
    Application.CheckLanguage = True
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
    Selection.Find.Execute replace:=wdReplaceAll
End Sub


Sub ForeColor1()
    Selection.Font.TextColor = 15773696 ' +++
End Sub


Sub DottedUnderline()
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
    Selection.Find.Execute replace:=wdReplaceAll

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
    Selection.Find.Execute replace:=wdReplaceAll

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
            .name = fontName
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
    Selection.Find.Execute replace:=wdReplaceAll

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
    Selection.Find.Execute replace:=wdReplaceAll

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
    Selection.Find.Execute replace:=wdReplaceAll
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
        .Text = " i?? "
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
    Selection.Find.Execute replace:=wdReplaceAll

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
    Selection.Find.Execute replace:=wdReplaceAll

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
    Selection.Find.Execute replace:=wdReplaceAll

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
    Selection.Find.Execute replace:=wdReplaceAll

    ' Clearing begin and end of document
    Selection.HomeKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
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
    Selection.Find.Execute replace:=wdReplaceAll

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
    Selection.Find.Execute replace:=wdReplaceAll
End Sub


Sub replace(oldValue As String, newValue As String)
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
    Selection.Find.Execute replace:=wdReplaceAll
End Sub

'
' Clear clipboard:
'
'Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function EmptyClipboard Lib "user32" () As Long
'Private Declare Function CloseClipboard Lib "user32" () As Long
'
'Private Function FuncClearClipboard()
'    OpenClipboard (0&)
'    EmptyClipboard
'    CloseClipboard
'End Function
'
'Sub ClearClipboard() ' Call this Sub, it uses previous arrangements
'    Call FuncClearClipboard
'End Sub


'
'
Sub CopyHyperlinkToClipboard()
' Add reference to %systemroot%\System32\FM20.dll
' using Tools | References...
    Dim result As String
    Dim Obj As New DataObject

    result = Selection.Hyperlinks(1).Address

    Obj.SetText result
    Obj.PutInClipboard
End Sub


'
'
Sub CopyAllHyperlinksToClipboard()
' Add reference to %systemroot%\System32\FM20.dll
' using Tools | References...
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


Sub AppendWorkFile(name)
' Used by AddAllWorkFiles().
'
    Selection.TypeText ">>>>"
    Selection.TypeParagraph
    Selection.TypeText name
    Selection.TypeParagraph
    Selection.InsertFile FileName:=name, Link:=False
    Selection.TypeParagraph
    Selection.TypeParagraph
End Sub


Sub AddAllWorkFiles()
' Run the Sub in new Word document.
    AppendWorkFile ("C:\Users\Path\file.name")
    AppendWorkFile ("C:\Users\Path\file2.name")
    
    MsgBox "Done!", vbInformation
End Sub


Sub TestForUnsavedChanges()
    If ActiveDocument.Saved = False Then
        MsgBox "This document contains unsaved changes.", vbInformation
    Else
        MsgBox "The document is saved.", vbInformation
    End If
End Sub


Sub SaveWithCheck()
    If ActiveDocument.Saved = False Then
        ActiveDocument.Save
    End If
    
    MsgBox "The file successfully saved.", vbInformation
End Sub


Sub SaveAllWithCheck()
' https://www.extendoffice.com/documents/excel/2971-excel-save-all-open-files.html
    Dim xWb As Document
    For Each xWb In Application.Documents
        If Not xWb.ReadOnly _
                And xWb.Saved = False _
        Then
            xWb.Save
        End If
    Next
    
    MsgBox "Files successfully saved.", vbInformation
End Sub


Sub FindByShadingColor()
    ' wdColorAutomatic
    ' 49407 – Gold; 5287936 – Green;
    ' -587137089 – Dark (Code)
    SelectNext (49407)
End Sub


Sub FindByShadingColorExclude()
    Dim selStart As Double
    Dim selEnd As Double
    Dim NewSelection As SelectionObject
    Dim ExcludeColor As Double
    
    ExcludeColor = wdColorAutomatic
    
    selStart = Selection.Start
    selEnd = Selection.End
    Set NewSelection = SelectNext2(ExcludeColor)
    
    Do While NewSelection.selStart = selEnd
        selStart = NewSelection.selStart
        selEnd = NewSelection.selEnd
        Set NewSelection = SelectNext2(ExcludeColor)
    Loop
    
    Selection.Start = selEnd
    Selection.End = NewSelection.selStart
End Sub


Function SelectNext(color As Double) As SelectionObject
    Dim SelectionStart As Double
    Dim SelectionEnd As Double
    Dim SelectionFont As SelectionObject
    Dim SelectionParagraph As SelectionObject
    
    SelectionStart = Selection.Start
    SelectionEnd = Selection.End
    
    Set SelectionFont = SelectNextFont(color)
    
    Selection.Start = SelectionStart
    Selection.End = SelectionEnd
    
    Set SelectionParagraph = SelectNextParagraph(color)
    
    If SelectionFont.selStart < SelectionParagraph.selStart _
            And SelectionFont.selEnd < SelectionParagraph.selEnd _
        Then
        If (SelectionFont.selStart > SelectionStart _
                And SelectionParagraph.selStart > SelectionStart) _
            Or (SelectionFont.selStart < SelectionStart _
                And SelectionParagraph.selStart < SelectionStart) _
            Then
            Set SelectNext = SelectionFont
        Else
            Set SelectNext = SelectionParagraph
        End If
    Else
        If (SelectionFont.selStart > SelectionStart _
                And SelectionParagraph.selStart > SelectionStart) _
            Or (SelectionFont.selStart < SelectionStart _
                And SelectionParagraph.selStart < SelectionStart) _
            Then
            Set SelectNext = SelectionParagraph
        Else
            Set SelectNext = SelectionFont
        End If
    End If
    
        Selection.Start = SelectNext.selStart
        Selection.End = SelectNext.selEnd
End Function


Function SelectNextFont(color As Double) As SelectionObject
    Dim result As SelectionObject
    Set result = New SelectionObject
    
    Selection.Find.ClearFormatting
    
    Selection.Find.Font.Shading.BackgroundPatternColor = color
    
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    
    result.selStart = Selection.Start
    result.selEnd = Selection.End
    
    Set SelectNextFont = result
End Function


Function SelectNextParagraph(color As Double) As SelectionObject
    Dim result As SelectionObject
    Set result = New SelectionObject
    
    Selection.Find.ClearFormatting
    
    Selection.Find.ParagraphFormat.Shading.BackgroundPatternColor = color
    
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    
    result.selStart = Selection.Start
    result.selEnd = Selection.End
    
    Set SelectNextParagraph = result
End Function


Function SelectNext2(color As Double) As SelectionObject
' https://answers.microsoft.com/en-us/msoffice/forum/msoffice_word-mso_other-mso_2007/find-shading/4c31b820-3457-453c-9b1c-672d41d7c013?auth=1
' Does not work for styled paragraphs – maybe it's need to configure some Selection.Find's options.
    Dim result As SelectionObject
    Set result = New SelectionObject
    
    Selection.Find.ClearFormatting
    '
    ' ** Toggle the next two options to find in text or entire paragraph
    'Selection.Find.Font.Shading.BackgroundPatternColor = color
    Selection.Find.ParagraphFormat.Shading.BackgroundPatternColor = color
    
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        '.Style = "Code"
    End With
    Selection.Find.Execute
    
    result.selStart = Selection.Start
    result.selEnd = Selection.End
    
    Set SelectNext2 = result
End Function


Sub ManageDocumentsHistory()
On Error GoTo TheError
    Dim RecentFilesStr As String
    Dim RecentFilesNum As Integer
    
    RecentFilesStr = InputBox("Enter recent files number (0-50):", , RecentFiles.Maximum)
    
    If RecentFilesStr = vbNullString Then
        Exit Sub
    End If
    
    RecentFilesNum = RecentFilesStr
    
    If RecentFilesNum < 0 Then
        RecentFilesNum = 0
    ElseIf RecentFilesNum > 50 Then
        RecentFilesNum = 50
    End If
    
    RecentFiles.Maximum = RecentFilesNum
    GoTo TheEnd
    
TheError:
    MsgBox "Input Error. Restart the macro and enter a correct number (0-50).", vbCritical
    Exit Sub
TheEnd:
End Sub


Sub ZoomTo()
On Error GoTo TheError
    Dim PercentageStr As String
    Dim PercentageNum As Integer
    
    PercentageStr = InputBox("Enter zoom percentage:", , ActiveWindow.ActivePane.View.Zoom.Percentage)
    
    If PercentageStr = vbNullString Then
        Exit Sub
    End If
    
    PercentageNum = PercentageStr
    
    If PercentageNum < 10 Then
        PercentageNum = 10
    ElseIf PercentageNum > 500 Then
        PercentageNum = 500
    End If
    
    ActiveWindow.ActivePane.View.Zoom.Percentage = PercentageNum
    GoTo TheEnd
    
TheError:
    MsgBox "Input Error. Restart the macro and enter a correct number (10-500).", vbCritical
    Exit Sub
TheEnd:
End Sub


Sub ShowSelectionLength()
    MsgBox "Selection length: " + CStr(Selection.End - Selection.Start) + ".", vbInformation
End Sub


Sub ExplorePath()
    Shell Environ("windir") & "\Explorer.exe " & ActiveDocument.Path, vbMaximizedFocus
End Sub


Sub LinkToFilesReminder()
    MsgBox "File | Info | Edit Links to Files", vbInformation
End Sub


Public Function CountChrInString(Expression As String, Character As String) As Long
' https://stackoverflow.com/questions/9260982/how-to-find-number-of-occurences-of-slash-from-a-strings
    Dim iResult As Long
    Dim sParts() As String

    sParts = Split(Expression, Character)

    iResult = UBound(sParts, 1)

    If (iResult = -1) Then
    iResult = 0
    End If

    CountChrInString = iResult

End Function


Sub ReplaceLineBreaksInSelection()
    Dim replace, replaceTo As String
    Dim selStart, selEnd, lenReplace, lenReplaceTo, replaceCount As Long
    
    'replace = "Roman"
    'replaceTo = "Romasyk"
    replace = "^l"
    replaceTo = "; "
    
    lenReplace = Len(replace) - 1
    lenReplaceTo = Len(replaceTo)
    
    replaceCount = CountChrInString(Selection.Text, "^l")
    
    MsgBox replaceCount
    
    selStart = Selection.Start
    selEnd = Selection.End
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = replace
        .Replacement.Text = replaceTo
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute replace:=wdReplaceAll
    
    Selection.Start = selStart
    Selection.End = selEnd + (lenReplaceTo - lenReplace) * replaceCount
End Sub


Sub ReplaceParagraphMarksInSelection()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "; "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute replace:=wdReplaceAll
End Sub


Sub TrimCellSpaces()
' https://superuser.com/questions/1028319/remove-trailing-whitespace-at-the-end-of-table-cells
' Fixed using combination of Left() and Trim().
' Add a reference to Microsoft VBScript Regular Expressions 5.5.
    Dim myRE As New regExp
    Dim itable As Table
    Dim C As Cell
    Dim l, count As Integer
    myRE.Pattern = "\s+(?!.*\w)"
    count = 0
    For Each itable In ActiveDocument.Tables
        For Each C In itable.Range.Cells
            l = Len(C.Range.Text)
            With myRE
                ' C.Range.Text = .Replace(C.Range.Text, "")
                C.Range.Text = Trim(Left(C.Range.Text, l - 2))
            End With
            If Len(C.Range.Text) <> l Then
                count = count + 1
            End If
        Next
    Next
    
    MsgBox "Done " + CStr(count) + " replacements."
End Sub

>>>>
' Put in a module.
' It uses the SelectionHandler class module defined below.
' Stop the macro manually if it is needed.
' Version 1.
Dim X As New SelectionHandler

Sub AAHandleSelection()
    Set X.appWord = Word.Application
End Sub

'
'
' Version 2.
'
Dim X As New SelectionHandler
Public SelectionHandlerIsOn As Boolean

Sub AAHandleSelection()
    Set X.appWord = Word.Application
    
    If Not SelectionHandlerIsOn Then
        X.Process = True
        SelectionHandlerIsOn = True
    Else
        X.Process = False
        SelectionHandlerIsOn = False
    End If
End Sub

>>>>
Option Explicit

Sub Macro1()
    Dim selStart, selEnd As Long
    selStart = Selection.Start
    selEnd = Selection.End
    MsgBox "Start: " + CStr(selStart) + ", End: " + CStr(selEnd)
End Sub

Sub PageDown()
'
' PageDown Macro
'
'
    Selection.GoToNext wdGoToPage
End Sub

Sub PageUp()
'
' PageUp Macro
'
'
    Selection.GoToPrevious wdGoToPage
End Sub

>>>> Class Module SelectionHandler; 
' Version 1.
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
 ' Workaround – for some reasons Trim() does not trim a paragraph mark.
 If EndsWith(Sel.Text, Chr(13)) And (Len(Sel.Text) > 1) Then
    Sel.End = Sel.End - 1
 End If

 If EndsWith(Sel.Text, " ") And (Len(Sel.Text) > 1) Then
    diff = Len(Sel.Text) - Len(Trim(Sel.Text))
    If diff <> Len(Sel.Text) Then
        Sel.End = Sel.End - diff
    End If
 End If
End Sub

'
'
' Version 2, toggle processing.
'
Option Explicit

' https://msdn.microsoft.com/en-us/vba/word-vba/articles/application-windowselectionchange-event-word
' https://msdn.microsoft.com/VBA/Word-VBA/articles/using-events-with-the-application-object-word
Public WithEvents appWord As Word.Application
Public Process As Boolean


' http://excelrevisited.blogspot.com/2012/06/endswith.html
Public Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     Dim ss As String
     endingLen = Len(ending)
     EndsWith = (Right(UCase(str), endingLen) = UCase(ending))
End Function

Private Sub appWord_WindowSelectionChange(ByVal Sel As Selection)
 If Not Process Then
    Exit Sub
 End If
 Dim diff As Integer
 ' Workaround – for some reasons Trim() does not trim a paragraph mark.
 If EndsWith(Sel.Text, Chr(13)) And (Len(Sel.Text) > 1) Then
    Sel.End = Sel.End - 1
 End If

 If EndsWith(Sel.Text, " ") And (Len(Sel.Text) > 1) Then
    diff = Len(Sel.Text) - Len(Trim(Sel.Text))
    If diff <> Len(Sel.Text) Then
        Sel.End = Sel.End - diff
    End If
 End If
End Sub


>>>> Module SelectionObject
Option Explicit

Public selStart As Double
Public selEnd As Double
