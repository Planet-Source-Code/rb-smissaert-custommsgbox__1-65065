Attribute VB_Name = "modMsgBox"
Option Explicit
Private Declare Function OleTranslateColor _
                          Lib "OLEPRO32.DLL" _
                              (ByVal OLE_COLOR As Long, _
                               ByVal HPALETTE As Long, _
                               pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Public bLog As Boolean

Function RemoveNonNumeric2(strString As String, _
                           Optional bIntegerOnly As Boolean, _
                           Optional bPositiveOnly As Boolean, _
                           Optional bClearAndOut As Boolean, _
                           Optional bHasChanged As Boolean) As String

  Dim i As Long
  Dim c As Long
  Dim bHasDot As Boolean
  Dim btArray() As Byte

  '45         -
  '46         .
  '48 to 58   0 to 9
  '------------------

  btArray = strString

  If bIntegerOnly Then
    If bPositiveOnly Then
      For i = 0 To UBound(btArray) Step 2
        If btArray(i) < 48 Or btArray(i) > 58 Then
          bHasChanged = True
          If bClearAndOut Then
            RemoveNonNumeric2 = ""
            Exit Function
          End If
          strString = Left$(strString, i \ 2 - c) & _
                      Mid$(strString, i \ 2 + 2 - c)
          c = c + 1
        End If
      Next i
    Else  'If bPositiveOnly
      For i = 0 To UBound(btArray) Step 2
        If btArray(i) < 48 Or btArray(i) > 58 Then
          If i - c * 2 = 0 Then
            If btArray(i) <> 45 Then
              bHasChanged = True
              If bClearAndOut Then
                RemoveNonNumeric2 = ""
                Exit Function
              End If
              strString = Left$(strString, i \ 2 - c) & _
                          Mid$(strString, i \ 2 + 2 - c)
              c = c + 1
            End If
          Else  'If i - c * 2 = 0
            bHasChanged = True
            If bClearAndOut Then
              RemoveNonNumeric2 = ""
              Exit Function
            End If
            strString = Left$(strString, i \ 2 - c) & _
                        Mid$(strString, i \ 2 + 2 - c)
            c = c + 1
          End If  'If i - c * 2 = 0
        End If
      Next i
    End If  'If bPositiveOnly
  Else  'If bIntegerOnly
    If bPositiveOnly Then
      For i = 0 To UBound(btArray) Step 2
        If btArray(i) < 48 Or btArray(i) > 58 Then
          If btArray(i) <> 46 Or bHasDot Then
            bHasChanged = True
            If bClearAndOut Then
              RemoveNonNumeric2 = ""
              Exit Function
            End If
            strString = Left$(strString, i \ 2 - c) & _
                        Mid$(strString, i \ 2 + 2 - c)
            c = c + 1
          Else
            bHasDot = True
          End If
        End If
      Next i
    Else  'If bPositiveOnly
      For i = 0 To UBound(btArray) Step 2
        If i - c * 2 = 0 Then
          If btArray(i) < 48 Or btArray(i) > 58 Then
            If btArray(i) <> 45 Then
              If btArray(i) <> 46 Or bHasDot Then
                bHasChanged = True
                If bClearAndOut Then
                  RemoveNonNumeric2 = ""
                  Exit Function
                End If
                strString = Left$(strString, i \ 2 - c) & _
                            Mid$(strString, i \ 2 + 2 - c)
                c = c + 1
              Else
                bHasDot = True
              End If
            End If
          End If
        Else  'If i - c * 2 = 0
          If btArray(i) < 48 Or btArray(i) > 58 Then
            If btArray(i) <> 46 Or bHasDot Then
              bHasChanged = True
              If bClearAndOut Then
                RemoveNonNumeric2 = ""
                Exit Function
              End If
              strString = Left$(strString, i \ 2 - c) & _
                          Mid$(strString, i \ 2 + 2 - c)
              c = c + 1
            Else
              bHasDot = True
            End If
          End If
        End If   'If i - c * 2 = 0
      Next i
    End If  'If bPositiveOnly
  End If  'If bIntegerOnly

  RemoveNonNumeric2 = strString

End Function

Function GetContrastingFontx(lClr As Long) As Long

  'returns a contrasting font colour, given a long colour
  '------------------------------------------------------
  Dim R As Long
  Dim G As Long
  Dim B As Long

  B = Int(lClr / 65536)
  G = Int((lClr Mod 65536) / 256)
  R = Int(lClr Mod 256)

  'calculation from Peter Thornton
  '-------------------------------
  If R * 0.206 + G * 0.679 + B * 0.115 > 135 Then
    GetContrastingFontx = vbBlack
  Else
    GetContrastingFontx = vbWhite
  End If

End Function

Function BreakText(strText As String, _
                   lMaxLen As Long, _
                   bIndent As Boolean, _
                   bIndentWithNumbersOnly As Boolean, _
                   strIndenter As String) As String

  Dim i As Long
  Dim strTemp As String
  Dim arr

  If InStr(1, strText, vbCrLf, vbBinaryCompare) > 0 Then
    arr = Split(strText, vbCrLf)
    For i = 0 To UBound(arr)
      arr(i) = BreakText2(arr(i), _
                          lMaxLen, _
                          bIndent, _
                          bIndentWithNumbersOnly, _
                          strIndenter)
    Next
    strTemp = arr(0)
    For i = 1 To UBound(arr)
      strTemp = strTemp & vbCrLf & arr(i)
    Next
  Else
    strTemp = BreakText2(strText, _
                         lMaxLen, _
                         bIndent, _
                         bIndentWithNumbersOnly, _
                         strIndenter)
  End If

  BreakText = strTemp

End Function

Function BreakText2(strText, _
                    lMaxLen As Long, _
                    bIndent As Boolean, _
                    bIndentWithNumbersOnly As Boolean, _
                    strIndenter As String) As String

  Dim i As Long
  Dim strTemp As String
  Dim strTest As String
  Dim strBefore As String
  Dim iBefore As Long
  Dim lLastBreak As Long
  Dim lTextWidthAim As Long
  Dim lTextWidth As Long
  Dim lTextWidthBefore As Long
  Dim lenText As Long
  Dim strIndenter2 As String
  Dim frm As FMsgBox

  On Error GoTo ERROROUT

  Set frm = New FMsgBox

  Do While frm.TextWidth(String(i, "a")) < Screen.Width - 2000
    i = i + 1
  Loop

  If i > lMaxLen Then
    i = lMaxLen
  End If

  lTextWidthAim = frm.TextWidth(String(i, "a"))
  lenText = Len(strText)
  strIndenter2 = strIndenter

  'this deals with doing an indent after a new linebreak
  '-----------------------------------------------------
  If bIndent Then
    If bIndentWithNumbersOnly Then
      If Asc(Left$(strText, 1)) < 48 Or _
         Asc(Left$(strText, 1)) > 57 Then
        strIndenter2 = ""
      End If
    End If
  Else
    strIndenter2 = ""
  End If

  For i = 1 To lenText

    If Mid$(strText, i, 1) = Chr(32) And (i - lLastBreak) > lMaxLen \ 2 Then

      strTest = Mid$(strText, lLastBreak + 1, i - (lLastBreak + 1))
      lTextWidth = frm.TextWidth(strTest)   'current TextWidth

      If lTextWidth < lTextWidthAim Then
        lTextWidthBefore = lTextWidth
        strBefore = strTest
        iBefore = i
      Else
        If lTextWidthAim - lTextWidthBefore < lTextWidth - lTextWidthAim Then
          lLastBreak = iBefore
          If Len(strTemp) = 0 Then
            strTemp = strBefore
          Else
            strTemp = strTemp & vbCrLf & strIndenter2 & strBefore
          End If
        Else
          lLastBreak = i
          If Len(strTemp) = 0 Then
            strTemp = strTest
          Else
            strTemp = strTemp & vbCrLf & strIndenter2 & strTest
          End If
        End If
      End If
    End If

  Next

  If lLastBreak - 1 < lenText Then
    If Len(strTemp) = 0 Then
      strTemp = strText
    Else
      strTemp = strTemp & vbCrLf & strIndenter2 & _
                Right$(strText, lenText - lLastBreak)
    End If
  End If

  BreakText2 = strTemp

  Unload frm
  Set frm = Nothing

  Exit Function
ERROROUT:

  BreakText2 = strTemp

  If Not frm Is Nothing Then
    Unload frm
    Set frm = Nothing
  End If

End Function

Function AddLineSpacer(strString As String, strLineSpacer As String) As String

  Dim arr
  Dim i As Long
  Dim strTemp As String

  If InStr(1, strString, vbCrLf, vbBinaryCompare) = 0 Then
    AddLineSpacer = strLineSpacer & strString
  Else
    arr = Split(strString, vbCrLf)
    For i = 0 To UBound(arr)
      arr(i) = strLineSpacer & arr(i)
    Next
    strTemp = arr(0)
    For i = 1 To UBound(arr)
      strTemp = strTemp & vbCrLf & arr(i)
    Next
    AddLineSpacer = strTemp
  End If

End Function

Function AutoSizeUnderlineX(ByVal strChar As String, _
                            ByVal strText As String) As String

  Dim i As Long
  Dim lLenText As Long
  Dim btTabWidthInSpaces As Byte
  Dim frm As FMsgBox

  On Error GoTo ERROROUT

  Set frm = New FMsgBox

  If InStr(1, strText, vbTab, vbBinaryCompare) > 0 Then
    btTabWidthInSpaces = frm.TextWidth(vbTab) \ frm.TextWidth(Chr(32))
    'without this replace the width of tabs is under-calculated
    '----------------------------------------------------------
    strText = Replace(strText, _
                      vbTab, _
                      String(btTabWidthInSpaces, Chr(32)), _
                      1, -1, _
                      vbBinaryCompare)
  End If

  i = 1
  lLenText = frm.TextWidth(strText)

  Do While frm.TextWidth(String(i, strChar)) < lLenText
    i = i + 1
  Loop

  AutoSizeUnderlineX = String(i, strChar)

  Unload frm
  Set frm = Nothing

  Exit Function
ERROROUT:

  'to give at least something back
  '-------------------------------
  AutoSizeUnderlineX = String(Len(strText), strChar)

  If Not frm Is Nothing Then
    Unload frm
    Set frm = Nothing
  End If

End Function

Function AdjustUnderlines(strString As String) As String

  Dim arr
  Dim i As Long
  Dim strTemp As String

  On Error GoTo ERROROUT

  'no linebreaks, get out
  '----------------------
  If InStr(1, strString, vbCrLf, vbBinaryCompare) = 0 Then
    AdjustUnderlines = strString
    Exit Function
  End If

  'no underlining, get out
  '-----------------------
  If InStr(1, strString, "----", vbBinaryCompare) = 0 And _
     InStr(1, strString, "____", vbBinaryCompare) = 0 Then
    AdjustUnderlines = strString
    Exit Function
  End If

  arr = Split(strString, vbCrLf)
  strTemp = arr(0)

  'adjust the underline strings
  '----------------------------
  For i = 1 To UBound(arr)
    If Left$(arr(i), 4) = "----" Or _
       Left$(arr(i), 4) = "____" Then
      arr(i) = AutoSizeUnderlineX(Left$(arr(i), 1), arr(i - 1))
    End If
    strTemp = strTemp & vbCrLf & arr(i)
  Next

  AdjustUnderlines = strTemp

  Exit Function
ERROROUT:

  'to give at least something back
  '-------------------------------
  AdjustUnderlines = strString

End Function

Function PadToWidthX(strStringOld As String, _
                     strAdd As String, _
                     strWidth As String, _
                     Optional bAddTab As Boolean = True) As String

  Dim strStringPrevious As String
  Dim siAimWidth As Single
  Dim siWidthPrevious As Single
  Dim siWidth As Single
  Dim strOriginal As String
  Dim frm As FMsgBox

  On Error GoTo ERROROUT

  Set frm = New FMsgBox

  strOriginal = strStringOld
  siAimWidth = frm.TextWidth(strWidth)
  siWidth = frm.TextWidth(strStringOld)

  Do While siWidth < siAimWidth
    siWidthPrevious = siWidth
    strStringPrevious = strStringOld
    strStringOld = strStringOld & strAdd
    siWidth = frm.TextWidth(strStringOld)
  Loop

  'to keep as close to the aimed width
  '-----------------------------------
  If siWidth > siAimWidth Then
    If siWidth - siAimWidth > _
       siAimWidth - siWidthPrevious Then
      strStringOld = strStringPrevious
    End If
  End If

  If bAddTab Then
    strStringOld = strStringOld & vbTab
  End If

  PadToWidthX = strStringOld

  Unload frm
  Set frm = Nothing

  Exit Function
ERROROUT:

  'give back the original if there was an error
  '--------------------------------------------
  PadToWidthX = strOriginal

  If Not frm Is Nothing Then
    Unload frm
    Set frm = Nothing
  End If

End Function

Function LineupTabsX(strOld As String, _
                     strSpacer As String, _
                     lGap As Long, _
                     Optional bAddTab As Boolean = True) As String

  Dim arr1
  Dim arr2
  Dim arr3
  Dim lCols As Long
  Dim lColsMax As Long
  Dim i As Long
  Dim c As Long
  Dim lWidth As Long
  Dim lMax As Long
  Dim strMax As String
  Dim strNew As String

  On Error GoTo ERROROUT

  arr1 = Split(strOld, vbCrLf)

  'get the column count for the array
  '----------------------------------
  For i = 0 To UBound(arr1)
    If InStr(1, arr1(i), vbTab, vbBinaryCompare) > 0 Then
      lCols = CountCharInStringX(CStr(arr1(i)), vbTab)
      If lCols > lColsMax Then
        lColsMax = lCols
      End If
    End If
  Next

  ReDim arr3(0 To UBound(arr1), 0 To lColsMax) As String

  'get the string in a 2-D array
  '-----------------------------
  For i = 0 To UBound(arr1)
    If InStr(1, arr1(i), vbTab, vbBinaryCompare) > 0 Then
      lCols = CountCharInStringX(CStr(arr1(i)), vbTab)
      For c = 0 To lCols
        arr2 = Split(arr1(i), vbTab)
        arr3(i, c) = arr2(c)
      Next
    Else
      'if there is no tab in the line only
      'fill the first element of the row
      '-----------------------------------
      arr3(i, 0) = arr1(i)
    End If
  Next

  'add the padding to line up the columns
  'no need to pad the last column !
  '--------------------------------------
  For c = 0 To lCols - 1
    For i = 0 To UBound(arr3)
      'only add padding if it was a row with tabs
      '------------------------------------------
      If Len(arr3(i, 1)) > 0 Then
        lWidth = Len(AutoSizeUnderlineX(strSpacer, CStr(arr3(i, c))))
        If lWidth > lMax Then
          'get the widest string for that column
          '-------------------------------------
          lMax = lWidth
          strMax = arr3(i, c)
        End If
      End If
    Next
    'pad according to the widest string in that column
    '-------------------------------------------------
    For i = 0 To UBound(arr3)
      'only add padding if it was a row with tabs
      '------------------------------------------
      If Len(arr3(i, 1)) > 0 Then
        arr3(i, c) = PadToWidthX(CStr(arr3(i, c)), _
                                 strSpacer, _
                                 strMax & String(lGap, strSpacer), _
                                 bAddTab)
      End If
    Next
    lMax = 0
  Next

  're-construct the new string
  '---------------------------
  For i = 0 To UBound(arr3)
    If Len(arr3(i, 1)) > 0 Then
      For c = 0 To lCols
        strNew = strNew & arr3(i, c)
      Next
    Else
      'no row with tabs, so just add the first element of the row
      '----------------------------------------------------------
      strNew = strNew & arr3(i, 0)
    End If
    If i < UBound(arr3) Then
      strNew = strNew & vbCrLf
    End If
  Next

  LineupTabsX = strNew

  Exit Function
ERROROUT:

  'give the old string back if there is an error
  '---------------------------------------------
  LineupTabsX = strOld

End Function

Function CountCharInStringX(strString As String, _
                            strChar As String, _
                            Optional bNotInQuotes As Boolean = True) As Long

  'working with a byte array is about twice
  'as fast compared to working with string functions
  '-------------------------------------------------
  Dim byteArray() As Byte
  Dim lAscChar As Byte
  Dim bInQuote As Boolean
  Dim i As Long
  Dim n As Long

  On Error GoTo ERROROUT

  byteArray = strString
  lAscChar = Asc(strChar)

  If bNotInQuotes Then
    For i = 0 To UBound(byteArray)
      If byteArray(i) = 34 Then
        bInQuote = (bInQuote = False)
      End If
      If bInQuote = False Then
        If byteArray(i) = lAscChar Then
          n = n + 1
        End If
      End If
    Next
  Else
    For i = 0 To UBound(byteArray)
      If byteArray(i) = lAscChar Then
        n = n + 1
      End If
    Next
  End If

  CountCharInStringX = n

  Exit Function
ERROROUT:

  CountCharInStringX = -1
  On Error GoTo 0

End Function

Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long

  ' Convert Automation color to Windows color
  If OleTranslateColor(oClr, hPal, TranslateColor) Then
    TranslateColor = CLR_INVALID
  End If

End Function

Public Function bFileExists(strFile As String) As Boolean

  Dim lAttr As Long

  On Error Resume Next
  lAttr = GetAttr(strFile)
  bFileExists = (Err.Number = 0) And ((lAttr And vbDirectory) = 0)
  On Error GoTo 0

End Function
