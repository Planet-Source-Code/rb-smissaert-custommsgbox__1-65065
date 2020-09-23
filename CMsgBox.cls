VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMsgBox"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function MsgBoxDLL(Optional vPrompt = "", Optional vTitle = "", _
                          Optional vButton1 = "OK", Optional vButton2 = "", _
                          Optional vButton3 = "", Optional vButton4 = "", _
                          Optional btDefault As Integer = 1, Optional lMessageIcon As Long = 0, _
                          Optional lFormColour As Long = -1, Optional lLabelColour As Long = -1, _
                          Optional lButtonColour As Long = -1, Optional lLabelFontColour As Long = -1, _
                          Optional lMaxLen As Long = 125, Optional bPromptBorder As Boolean = False, _
                          Optional bIndentAfterBreak As Boolean = False, Optional bOnlyIndentWithNumericFirstChar As Boolean = True, _
                          Optional strIndent As String, Optional bStartLineWithSpacer As Boolean = False, _
                          Optional strSpacer As String, Optional siPromptButtonFontSize As Single = 10, _
                          Optional lMinButtonWidth As Long = 1200, Optional lMinButtonGap As Long = 150, _
                          Optional lButtonEdge As Long = 150, Optional bSpaceButtons As Boolean = False, _
                          Optional bAdjustUnderlineStrings As Boolean = False, Optional bShowTitleIcon As Boolean = True, _
                          Optional bCenterPrompt As Boolean = False, _
                          Optional bInput As Boolean = False, Optional vInputDefault As Variant = "", _
                          Optional bNumericInput As Boolean = False, Optional bIntegerOnly As Boolean = False, _
                          Optional vMinValue As Variant = "", Optional vMaxValue As Variant = "", _
                          Optional bBeep As Boolean = False, Optional bLineUpTabs As Boolean = False, _
                          Optional strTabSpacer As String = " ", Optional lTabGap As Long = 1, _
                          Optional bAddTab As Boolean = True, Optional vToolTip1 = "", Optional vToolTip2 = "", _
                          Optional vToolTip3 = "", Optional vToolTip4 = "", Optional lActiveButtonColour As Long = -1, _
                          Optional bReturnLineNumber As Boolean, Optional bMaskInput As Boolean) As String

  '--------------------------------------------------------------------------'
  '45 optional arguments, returning a string                                 '
  'returns the caption of the pressed button or                              '
  'if it is an inputbox the text of the input textbox                        '
  'or if bReturnLineNumber = True the right-clicked linenumber of the prompt '
  '--------------------------------------------------------------------------'
  Dim frm As FMsgBox
  Dim i As Integer
  Dim n As Integer
  Dim oExcel As Object
  Dim lCursor As Long
  Dim btButtonCount As Byte
  Dim strPrompt As String
  Dim strTitle As String
  Dim arrStringCaptions(0 To 3) As String
  Dim arrBShortCuts(0 To 3) As Boolean
  Dim arrStringToolTips(0 To 3) As String
  Dim bSetExcelCursor As Boolean
  Dim lAmpersandPosI As Long
  Dim lAmpersandPosN As Long

  On Error Resume Next

  If oExcel Is Nothing Then
    Set oExcel = GetObject(, "Excel.Application")
    If Not oExcel Is Nothing Then
      lCursor = oExcel.cursor
    End If
  End If

  If bBeep Then
    Beep
  End If

  If bInput Then
    'to keep in line with a regular InputBox
    '---------------------------------------
    lMessageIcon = 0
    vButton1 = "OK"
    vButton2 = "Cancel"
    vButton3 = ""
    vButton4 = ""
  End If

  'to move buttons to the left
  '---------------------------
  If Len(vButton2) = 0 Then
    If Len(vButton3) > 0 Then
      vButton2 = vButton3
      vButton3 = ""
    Else
      If Len(vButton4) > 0 Then
        vButton2 = vButton4
        vButton4 = ""
      End If
    End If
  End If

  If Len(vButton3) = 0 Then
    If Len(vButton4) > 0 Then
      vButton3 = vButton4
      vButton4 = ""
    End If
  End If

  btButtonCount = 1
  arrStringCaptions(0) = CStr(vButton1)
  arrStringToolTips(0) = CStr(vToolTip1)

  If Len(vButton2) > 0 Then
    btButtonCount = 2
    arrStringCaptions(1) = CStr(vButton2)
    arrStringToolTips(1) = CStr(vToolTip2)
  End If

  If Len(vButton3) > 0 Then
    btButtonCount = 3
    arrStringCaptions(2) = CStr(vButton3)
    arrStringToolTips(2) = CStr(vToolTip3)
  End If

  If Len(vButton4) > 0 Then
    btButtonCount = 4
    arrStringCaptions(3) = CStr(vButton4)
    arrStringToolTips(3) = CStr(vToolTip4)
  End If

  'to avoid making a shortcut underline if
  'another button starts with the same character
  '---------------------------------------------
  If btButtonCount > 1 And bInput = False Then
    For i = 0 To btButtonCount - 1
      arrBShortCuts(i) = True
      lAmpersandPosI = InStr(1, arrStringCaptions(i), "&", vbBinaryCompare)
      For n = 0 To btButtonCount - 1
        If i <> n Then
          lAmpersandPosN = InStr(1, arrStringCaptions(n), "&", vbBinaryCompare)
          If LCase(Mid$(arrStringCaptions(i), lAmpersandPosI + 1, 1)) = _
             LCase(Mid$(arrStringCaptions(n), lAmpersandPosN + 1, 1)) Then
            arrBShortCuts(i) = False
            arrBShortCuts(n) = False
            arrStringCaptions(i) = Replace(arrStringCaptions(i), "&", "")
            arrStringCaptions(n) = Replace(arrStringCaptions(n), "&", "")
            Exit For
          End If
        End If
      Next
    Next

    For i = 0 To btButtonCount - 1
      If arrBShortCuts(i) = True And _
         InStr(1, arrStringCaptions(i), "&", vbBinaryCompare) = 0 Then
        arrStringCaptions(i) = "&" & arrStringCaptions(i)
      End If
    Next
  End If

  strPrompt = CStr(vPrompt)
  strTitle = CStr(vTitle)

  If bLineUpTabs Then
    strPrompt = LineupTabsX(strPrompt, strTabSpacer, lTabGap, bAddTab)
  End If

  If bAdjustUnderlineStrings Then
    strPrompt = AdjustUnderlines(strPrompt)
  End If

  If bStartLineWithSpacer Then
    lMaxLen = lMaxLen - Len(strSpacer)
  End If

  If Len(strPrompt) > lMaxLen Then
    strPrompt = BreakText(strPrompt, _
                          lMaxLen, bIndentAfterBreak, _
                          bOnlyIndentWithNumericFirstChar, _
                          strIndent)
  End If

  If bStartLineWithSpacer Then
    strPrompt = AddLineSpacer(strPrompt, strSpacer)
  End If

  If lFormColour = -1 Then
    'to get the default long form colour
    lFormColour = TranslateColor(vbButtonFace)
  End If

  'so setting the form colour only will
  'give the label and buttons the same colour
  '------------------------------------------
  If lFormColour > -1 Then
    If lLabelColour = -1 Then
      lLabelColour = lFormColour
    End If
    If lButtonColour = -1 Then
      lButtonColour = lFormColour
    End If
  End If

  'to avoid a non-contrasting label font colour
  '--------------------------------------------
  If lLabelColour > -1 Then
    If lLabelFontColour = -1 Then
      lLabelFontColour = GetContrastingFontx(lLabelColour)
    End If
  End If

  If Forms.Count = 0 Then
    Set frm = New FMsgBox
    bSetExcelCursor = True
  Else
    Set frm = New FMsgBox
    frm.SecondForm = True
  End If

  With frm
    .ButtonCount = btButtonCount
    .Prompt = strPrompt
    .Title = strTitle

    .cmdButton(0).ToolTipText = arrStringToolTips(0)

    'add more buttons if needed
    '--------------------------
    If btButtonCount > 1 Then
      For i = 2 To btButtonCount
        Load .cmdButton(i - 1)
        .cmdButton(i - 1).ToolTipText = _
        arrStringToolTips(i - 1)
      Next
    End If

    'add an input textbox if needed
    '------------------------------
    If bInput Then
      Load .txtPrompt(1)
    End If

    For i = 0 To btButtonCount - 1
      .ButtonCaptions(i) = arrStringCaptions(i)
    Next

    Set .objExcel = oExcel

    .DoInput = bInput
    .ReturnLineNumber = bReturnLineNumber
    .InputDefault = vInputDefault
    .NumericInput = bNumericInput
    .IntegerOnly = bIntegerOnly
    .MinValue = vMinValue
    .MaxValue = vMaxValue
    .DefaultButton = btDefault
    .MessageIcon = lMessageIcon
    .FormColour = lFormColour
    .LabelColour = lLabelColour
    .ButtonColour = lButtonColour
    .ActiveButtonColour = lActiveButtonColour
    .LabelFontColour = lLabelFontColour
    .MaxLenPrompt = lMaxLen
    .PromptBorder = bPromptBorder
    .PromptButtonFontSize = siPromptButtonFontSize
    .MinButtonWidth = lMinButtonWidth
    .MinButtonGap = lMinButtonGap
    .ButtonEdge = lButtonEdge
    .SpaceButtons = bSpaceButtons
    .ShowTitleIcon = bShowTitleIcon
    .CenterPrompt = bCenterPrompt
    .MaskInput = bMaskInput

    .PreActivate
    .Show vbModal

    MsgBoxDLL = .ReturnValue
  End With

  Unload frm
  Set frm = Nothing

  On Error Resume Next

  If bSetExcelCursor Then  'don't do cursor if the first form is still there
    If Not oExcel Is Nothing Then
      oExcel.cursor = lCursor
      Set oExcel = Nothing
    End If
  End If

End Function

'PT, do these need to be visible to externally, if not change Public > Friend
Public Function AutoSizeUnderline(strChar As String, _
                                  strStringCompare As String) As String

  AutoSizeUnderline = AutoSizeUnderlineX(strChar, strStringCompare)

End Function

Public Function PadToWidth(strStringOld As String, _
                           strAdd As String, _
                           strWidth As String, _
                           Optional bAddTab As Boolean = True) As String

  PadToWidth = PadToWidthX(strStringOld, _
                           strAdd, _
                           strWidth, _
                           bAddTab)

End Function

Public Function LineupTabs(strOld As String, _
                           strSpacer As String, _
                           lGap As Long, _
                           Optional bAddTab As Boolean = True) As String

  LineupTabs = LineupTabsX(strOld, _
                           strSpacer, _
                           lGap, _
                           bAddTab)
End Function

Public Function CountCharInString(strString As String, _
                                  strChar As String, _
                                  Optional bNotInQuotes As Boolean = True) _
                                  As Long
  CountCharInString = CountCharInStringX(strString, _
                                         strChar, _
                                         bNotInQuotes)
End Function

Public Function RemoveNonNumeric(strString As String, _
                                 Optional bIntegerOnly As Boolean, _
                                 Optional bPositiveOnly As Boolean, _
                                 Optional bClearAndOut As Boolean) As String
  RemoveNonNumeric = RemoveNonNumeric2(strString, _
                                       bIntegerOnly, _
                                       bPositiveOnly, _
                                       bClearAndOut)
End Function
