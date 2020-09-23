VERSION 5.00
Begin VB.Form FMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   1635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox cmdButton 
      Height          =   375
      Index           =   0
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   585
      Width           =   720
   End
   Begin VB.TextBox txtPrompt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "FMsgBox.frx":058A
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000

Private Const WM_LBUTTONDOWN = &H201
Private Declare Function SendMessage _
                          Lib "user32" _
                              Alias "SendMessageA" _
                              (ByVal hwnd As Long, _
                               ByVal wMsg As Long, _
                               ByVal wParam As Long, _
                               lParam As Any) As Long

Private Const EM_CHARFROMPOS = &HD7
Private Const EM_LINEFROMCHAR = &HC9

Private Declare Function DrawMenuBar Lib "user32" _
                                     (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" _
                                          (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" _
                                       (ByVal hwnd As Long, _
                                        ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" _
                                    (ByVal hMenu As Long, _
                                     ByVal nPosition As Long, _
                                     ByVal wFlags As Long) _
                                     As Long
Private Declare Function LoadStandardIcon Lib "user32" _
                                          Alias "LoadIconA" _
                                          (ByVal hInstance As Long, _
                                           ByVal lpIconNum As enStandardIconEnum) _
                                           As Long
Private Declare Function DrawIcon Lib "user32" _
                                  (ByVal hdc As Long, _
                                   ByVal x As Long, _
                                   ByVal y As Long, _
                                   ByVal hIcon As Long) As Long
Private Enum enStandardIconEnum
  IDI_APPLICATION = 32512&
  IDI_ASTERISK = 32516&
  IDI_EXCLAMATION = 32515&
  IDI_HAND = 32513&
  IDI_ERROR = IDI_HAND
  IDI_INFORMATION = IDI_ASTERISK
  IDI_QUESTION = 32514&
  IDI_WARNING = IDI_EXCLAMATION
  IDI_WINLOGO = 32517
End Enum

Private Const SB_HORZ As Long = 0
Private Const SB_VERT As Long = 1
Private Const SB_BOTH As Long = 3
Private Declare Function ShowScrollBar Lib "user32" _
                                       (ByVal hwnd As Long, _
                                        ByVal wBar As Long, _
                                        ByVal bShow As Long) As Long

Private m_Excel As Object
Private m_bSecondForm As Boolean
Private m_DoInput As Boolean
Private m_ReturnLineNumber As Boolean
Private m_InputDefault As String
Private m_InputNumeric As Boolean
Private m_IntegerOnly As Boolean
Private m_MinValue As Variant
Private m_MaxValue As Variant
Private m_ButtonCount As Byte
Private m_Icon As Long
Private m_Prompt As String
Private m_Title As String
Private m_ButtonCaptions(0 To 3) As String
Private m_Default As Integer
Private m_ColourForm As Long
Private m_ColourLabel As Long
Private m_ColourButton As Long
Private m_ColourLabelFont As Long
Private m_ColourActiveButton As Long
Private m_MaxLen As Long
Private m_PromptBorder As Boolean
Private m_PromptButtonFontSize As Single
Private m_MinButtonWidth As Long
Private m_MinButtonGap As Long
Private m_ButtonEdge As Long
Private m_SpaceButtons As Boolean
Private m_ShowTitleIcon As Boolean
Private m_CenterPrompt As Boolean
Private m_MaskInput As Boolean

Private lFormWidth As Long
Private lPromptHeight As Long
Private lInputHeight As Long
Private s_RetVal As String
Private bCanEscape As Boolean
Private bReadyForInput As Boolean
Private bDoneCursor As Boolean
Private btActiveButtonIndex As Byte
Private lButtonFontColour As Long
Private lActiveButtonFontColour As Long
Private lCursorPos As Long
Private arrButtonResults() As String

Public Property Set objExcel(oExcel As Object)
  Set m_Excel = oExcel
End Property

Public Property Let SecondForm(ByVal bValue As Boolean)
  m_bSecondForm = bValue
End Property
Public Property Get SecondForm() As Boolean
  SecondForm = m_bSecondForm
End Property

Public Property Let DoInput(ByVal bValue As Boolean)
  m_DoInput = bValue
End Property

Public Property Let ReturnLineNumber(ByVal bValue As Boolean)
  m_ReturnLineNumber = bValue
End Property

Public Property Let NumericInput(ByVal bValue As Boolean)
  m_InputNumeric = bValue
End Property

Public Property Let IntegerOnly(ByVal bValue As Boolean)
  m_IntegerOnly = bValue
End Property

Public Property Let InputDefault(ByVal vValue As Variant)
  m_InputDefault = vValue
End Property

Public Property Let MinValue(ByVal vValue As Variant)
  m_MinValue = vValue
End Property

Public Property Let MaxValue(ByVal vValue As Variant)
  m_MaxValue = vValue
End Property

Public Property Let ButtonCount(ByVal btValue As Byte)
  m_ButtonCount = btValue
End Property

Public Property Let MessageIcon(ByVal lValue As Long)
  m_Icon = lValue
End Property

Public Property Let Prompt(ByVal sText As String)
  m_Prompt = sText
End Property

Public Property Let Title(ByVal sText As String)
  m_Title = sText
End Property

Public Property Let MaskInput(ByVal bValue As Boolean)
  m_MaskInput = bValue
End Property

Property Let ButtonCaptions(ByVal Index As Integer, sText As String)
  m_ButtonCaptions(Index) = sText
End Property

Public Property Let DefaultButton(ByVal iValue As Integer)
  m_Default = iValue
End Property

Public Property Let FormColour(ByVal lValue As Long)
  m_ColourForm = lValue
End Property

Public Property Let LabelColour(ByVal lValue As Long)
  m_ColourLabel = lValue
End Property

Public Property Let ButtonColour(ByVal lValue As Long)
  m_ColourButton = lValue
End Property

Public Property Let ActiveButtonColour(ByVal lValue As Long)
  m_ColourActiveButton = lValue
End Property

Public Property Let LabelFontColour(ByVal lValue As Long)
  m_ColourLabelFont = lValue
End Property

Public Property Let MaxLenPrompt(ByVal lValue As Long)
  m_MaxLen = lValue
End Property

Public Property Let PromptBorder(bValue As Boolean)
  m_PromptBorder = bValue
End Property

Public Property Let MinButtonWidth(ByVal lValue As Long)
  m_MinButtonWidth = lValue
End Property

Public Property Let MinButtonGap(ByVal lValue As Long)
  m_MinButtonGap = lValue
End Property

Public Property Let ButtonEdge(ByVal lValue As Long)
  m_ButtonEdge = lValue
End Property

Public Property Let SpaceButtons(ByVal bValue As Boolean)
  m_SpaceButtons = bValue
End Property

Public Property Let PromptButtonFontSize(ByVal siValue As Single)
  m_PromptButtonFontSize = siValue
End Property

Public Property Let ShowTitleIcon(ByVal bValue As Boolean)
  m_ShowTitleIcon = bValue
End Property

Public Property Let CenterPrompt(ByVal bValue As Boolean)
  m_CenterPrompt = bValue
End Property

Public Property Let ReturnValue(ByVal sValue As String)
  s_RetVal = sValue
End Property
Public Property Get ReturnValue() As String
  ReturnValue = s_RetVal
End Property

Private Sub cmdButton_Click(Index As Integer)
  'm_ButtonCaptions will have the & already taken off
  'this happens directly after setting the button captions
  '-------------------------------------------------------
  If m_DoInput And Index <> 1 Then
    s_RetVal = txtPrompt(1)
    If m_InputNumeric Then
      s_RetVal = Val(s_RetVal)  'does this need error handler ?
    End If
  Else
    s_RetVal = arrButtonResults(Index)
  End If

  Me.Hide

End Sub

Private Sub txtPrompt_Validate(Index As Integer, Cancel As Boolean)  'PT val
  If Index = 1 And m_InputNumeric Then
    Cancel = ValidateInput(txtPrompt(1))
  End If
End Sub

Private Function ValidateInput(sText As String) As Boolean  'PT val

  Dim cmb As CMsgBox
  Dim sPrompt As String
  Dim sTitle As String
  Dim vVal

  ' error handler ?
  If m_InputNumeric Then

    vVal = Val(sText)
    sTitle = "input check"

    If Len(m_MinValue) > 0 Then
      If vVal < Val(m_MinValue) Then
        sPrompt = "Input value too small" & vbCrLf & vbCrLf & _
                  "Needs to be at least " & m_MinValue
      End If
    End If

    If Len(sPrompt) = 0 Then
      If Len(m_MaxValue) > 0 Then
        If vVal > Val(m_MaxValue) Then
          sPrompt = "Input value too big" & vbCrLf & vbCrLf & _
                    "Can't be more than " & m_MaxValue
        End If
      End If
    End If

  End If

  If Len(sPrompt) Then
    Set cmb = New CMsgBox
    cmb.MsgBoxDLL vPrompt:=sPrompt, _
                  vTitle:=sTitle, _
                  lMessageIcon:=3, _
                  lFormColour:=m_ColourForm
    ValidateInput = True
  End If

End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cmdButton_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyDown KeyCode
  DoEvents
  KeyCode = 0
End Sub

Private Sub cmdButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

  KeyDown KeyCode
  DoEvents
  KeyCode = 0

End Sub

Private Sub txtPrompt_KeyDown(Index As Integer, _
                              KeyCode As Integer, _
                              Shift As Integer)

  If Index = 0 Then
    Exit Sub
  End If

  Select Case KeyCode
    Case 13
      KeyCode = 0
      If m_InputNumeric Then  'PT val
        If ValidateInput(Me.txtPrompt(1).Text) Then Exit Sub
      End If

      cmdButton_Click 0    'OK button

    Case 27
      KeyCode = 0
      cmdButton_Click 1    'Cancel button
  End Select

End Sub

Private Sub txtPrompt_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 1 Then
    lCursorPos = txtPrompt(1).SelStart
  End If
End Sub

Private Sub txtPrompt_Change(Index As Integer)

  Dim strTemp As String
  Dim bHasChanged As Boolean

  If Index = 0 Or bReadyForInput = False Or m_InputNumeric = False Then
    Exit Sub
  End If

  strTemp = RemoveNonNumeric2(txtPrompt(1).Text, _
                              m_IntegerOnly, _
                              Len(m_MinValue) > 0 And _
                              Val(m_MinValue) >= 0, _
                              False, _
                              bHasChanged)

  If bHasChanged Then
    txtPrompt(1).Text = strTemp
    If lCursorPos <= Len(txtPrompt(1).Text) Then
      txtPrompt(1).SelStart = lCursorPos
    End If
  End If

End Sub

Private Sub txtPrompt_MouseDown(Index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

  Dim lHwnd As Long
  Dim lPoints As Long
  Dim lResult As Long
  Dim lCharPos As Long
  Dim lLineNumber As Long
  Dim x2 As Long
  Dim y2 As Long

  If Index = 0 And m_ReturnLineNumber And Button = 2 Then

    'to avoid the right-mouse popup
    '------------------------------
    SendMessage Me.hwnd, WM_LBUTTONDOWN, 0, 0

    x2 = ScaleX(x, vbTwips, vbPixels)
    y2 = ScaleY(y, vbTwips, vbPixels)

    lHwnd = txtPrompt(0).hwnd
    lPoints = y2 * &H10000 + x2

    lResult = SendMessage(lHwnd, EM_CHARFROMPOS, 0&, ByVal lPoints)
    'lCharPos = LoWord(lResult)
    lLineNumber = HiWord(lResult)
    s_RetVal = lLineNumber + 1
    Me.Hide
    
  End If

End Sub

Private Function LoWord(DWord As Long) As Integer
  If DWord And &H8000& Then
    LoWord = &H8000 Or (DWord And &H7FFF&)
  Else
    LoWord = DWord And &HFFFF&
  End If
End Function

Private Function HiWord(DWord As Long) As Integer
  If DWord And &H80000000 Then
    HiWord = (DWord \ 65535) - 1
  Else
    HiWord = DWord \ 65535
  End If
End Function

Private Sub KeyDown(iKeyCode As Integer)

  Dim i As Integer
  Dim lAmpersandPos As Long
  Dim iClick As Integer
  Dim bClick As Boolean
  Dim btCount As Byte

  Select Case iKeyCode
    Case 13, 32
      For i = 0 To m_ButtonCount - 1
        If cmdButton(i) Is ActiveControl Then
          cmdButton_Click i
          DoEvents
          Exit For
        End If
      Next
    Case 27
      If bCanEscape Then
        If m_ButtonCount = 1 Then
          cmdButton_Click 0
          DoEvents
        Else
          For i = 1 To 3
            If LCase$(arrButtonResults(i)) = "cancel" Then
              cmdButton_Click i
              DoEvents
              Exit For
            End If
          Next
        End If
      End If
    Case Else
      If iKeyCode > 95 And iKeyCode < 106 Then
        'to get the non-numeric number key
        iKeyCode = iKeyCode - 48
      End If
      'if there is only one button the regular MsgBox
      'doesn't do keyboard letter shortcuts
      '----------------------------------------------
      If m_ButtonCount > 1 Then
        For i = 0 To m_ButtonCount - 1
          lAmpersandPos = InStr(1, m_ButtonCaptions(i), "&", vbBinaryCompare)
          If lAmpersandPos = 0 Then
            If iKeyCode = Asc(LCase(Left$(m_ButtonCaptions(i), 1))) Or _
               iKeyCode = Asc(UCase(Left$(m_ButtonCaptions(i), 1))) Then
              btCount = btCount + 1
              iClick = i
              bClick = True
              If btCount > 1 Then
                'as we wouldn't know what button to click
                '----------------------------------------
                Exit Sub
              End If
            End If
          Else
            If iKeyCode = Asc(LCase(Mid$(m_ButtonCaptions(i), lAmpersandPos + 1, 1))) Or _
               iKeyCode = Asc(UCase(Mid$(m_ButtonCaptions(i), lAmpersandPos + 1, 1))) Then
              btCount = btCount + 1
              iClick = i
              bClick = True
              If btCount > 1 Then
                'as we wouldn't know what button to click
                '----------------------------------------
                Exit Sub
              End If
            End If
          End If
        Next
        If bClick Then
          cmdButton_Click iClick
          DoEvents
        End If
      End If
  End Select

End Sub

Private Sub cmdButton_MouseMove(Index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

        Dim lBreakPos As Long

10      On Error GoTo ERROROUT

20      If m_ColourActiveButton > -1 Then
30        If btActiveButtonIndex > 0 And btActiveButtonIndex - 1 <> Index Then
            'reset previous active button
40          cmdButton(btActiveButtonIndex - 1).BackColor = m_ColourButton
50          cmdButton(btActiveButtonIndex - 1).ForeColor = lButtonFontColour
60        End If

70        If btActiveButtonIndex - 1 <> Index Then
80          btActiveButtonIndex = Index + 1
90          cmdButton(Index).BackColor = m_ColourActiveButton
100         cmdButton(Index).ForeColor = lActiveButtonFontColour
110       End If
120     End If

130     If bDoneCursor = False Then
140       If Not m_Excel Is Nothing Then
150         m_Excel.cursor = -4143
160       End If
170       bDoneCursor = True
180     End If

190     Exit Sub
ERROROUT:

200     lBreakPos = InStr(1, m_Prompt, vbCrLf, vbBinaryCompare)
210     If lBreakPos = 0 Then
220       lBreakPos = Len(m_Prompt)
230     End If

240     MsgBox "Private Sub cmdButton_MouseMove" & vbCrLf & _
               "Line: " & Erl & vbCrLf & _
               Err.Description & vbCrLf & _
               "Error number: " & Err.Number & vbCrLf & vbCrLf & _
               "First line of message prompt:" & vbCrLf & _
               Left$(m_Prompt, lBreakPos)

End Sub

Private Sub txtPrompt_MouseMove(Index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

  Dim lBreakPos As Long

  On Error GoTo ERROROUT

  If btActiveButtonIndex > 0 And m_ColourActiveButton > -1 Then
    'reset previous active button
    cmdButton(btActiveButtonIndex - 1).BackColor = m_ColourButton
    cmdButton(btActiveButtonIndex - 1).ForeColor = lButtonFontColour
    btActiveButtonIndex = 0
  End If

  If bDoneCursor = False Then
    If Not m_Excel Is Nothing Then
      m_Excel.cursor = -4143
    End If
    bDoneCursor = True
  End If

  Exit Sub
ERROROUT:

  lBreakPos = InStr(1, m_Prompt, vbCrLf, vbBinaryCompare)
  If lBreakPos = 0 Then
    lBreakPos = Len(m_Prompt)
  End If

  MsgBox "Private Sub txtPrompt_MouseMove" & vbCrLf & _
         "Line: " & Erl & vbCrLf & _
         Err.Description & vbCrLf & _
         "Error number: " & Err.Number & vbCrLf & vbCrLf & _
         "First line of message prompt:" & vbCrLf & _
         Left$(m_Prompt, lBreakPos)

End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)

        Dim lBreakPos As Long

10      On Error GoTo ERROROUT

20      If btActiveButtonIndex > 0 And m_ColourActiveButton > -1 Then
          'reset previous active button
30        cmdButton(btActiveButtonIndex - 1).BackColor = m_ColourButton
40        cmdButton(btActiveButtonIndex - 1).ForeColor = lButtonFontColour
50        btActiveButtonIndex = 0
60      End If

70      If bDoneCursor = False Then
80        If Not m_Excel Is Nothing Then
90          m_Excel.cursor = -4143
100       End If
110       bDoneCursor = True
120     End If

130     Exit Sub
ERROROUT:

140     lBreakPos = InStr(1, m_Prompt, vbCrLf, vbBinaryCompare)
150     If lBreakPos = 0 Then
160       lBreakPos = Len(m_Prompt)
170     End If

180     MsgBox "Private Sub Form_MouseMove" & vbCrLf & _
               "Line: " & Erl & vbCrLf & _
               Err.Description & vbCrLf & _
               "Error number: " & Err.Number & vbCrLf & vbCrLf & _
               "First line of message prompt:" & vbCrLf & _
               Left$(m_Prompt, lBreakPos)

End Sub

Private Sub Form_Paint()
  lPromptHeight = TextHeight(m_Prompt)
  Select Case m_Icon
    Case 1
      DisplayFormIcon IDI_ASTERISK
    Case 2
      DisplayFormIcon IDI_QUESTION
    Case 3
      DisplayFormIcon IDI_EXCLAMATION
    Case 4
      DisplayFormIcon IDI_HAND
  End Select
End Sub

Private Sub DisplayFormIcon(IconType As enStandardIconEnum)

        Dim lMidPrompt As Long
        Dim lBreakPos As Long

10      On Error GoTo ERROROUT

20      Me.Cls

30      If m_PromptBorder Then
40        lMidPrompt = (120 + 36 + lPromptHeight / 2) - 256
50      Else
60        lMidPrompt = (120 + lPromptHeight / 2) - 256
70      End If

        'hexadecimal 10& will be 16 in decimal
        '-------------------------------------
80      If lPromptHeight < 700 Then
90        DrawIcon Me.hdc, 10, 10&, LoadStandardIcon(0&, IconType)
100     Else
110       DrawIcon Me.hdc, _
                   10, _
                   CLng(Me.ScaleY(lMidPrompt, vbTwips, vbPixels)), _
                   LoadStandardIcon(0&, IconType)
120     End If

130     Exit Sub
ERROROUT:

140     lBreakPos = InStr(1, m_Prompt, vbCrLf, vbBinaryCompare)
150     If lBreakPos = 0 Then
160       lBreakPos = Len(m_Prompt)
170     End If

180     MsgBox "DisplayFormIcon" & vbCrLf & _
               "Line: " & Erl & vbCrLf & _
               Err.Description & vbCrLf & _
               "Error number: " & Err.Number & vbCrLf & vbCrLf & _
               "First line of message prompt:" & vbCrLf & _
               Left$(m_Prompt, lBreakPos)

End Sub

Sub DisableCloseButton(hwnd As Long)

  Dim hMenu As Long
  Dim menuItemCount As Long

  hMenu = GetSystemMenu(hwnd, 0)

  If hMenu Then
    menuItemCount = GetMenuItemCount(hMenu)
    RemoveMenu hMenu, _
               menuItemCount - 1, _
               MF_REMOVE Or MF_BYPOSITION
    DrawMenuBar hwnd
  End If

End Sub

Private Sub Form_Activate()

  If m_DoInput = False Then
    cmdButton(m_Default - 1).SetFocus
  Else
    With txtPrompt(1)
      ShowScrollBar .hwnd, SB_VERT, False
      .TabStop = True
      .TabIndex = 0
      .SetFocus
    End With
  End If

  ShowScrollBar txtPrompt(0).hwnd, SB_VERT, False

End Sub

Public Sub PreActivate()  'PT

        Dim i As Integer
        Dim lExtraHeight As Long
        Dim lExtraWidth As Long
        Dim lMidForm As Long
        Dim lCaptionWidth As Long
        Dim lPromptWidth As Long
        Dim lPromptHeight As Long
        Dim lInputWidth As Long
        Dim lFormHeight As Long
        Dim lIconsCorrection As Long
        Dim lPromptWidthCorrection As Long
        Dim btTabWidthInSpaces As Byte
        Dim lPromptHeightCorrection As Long
        Dim lMaxButtonWidth As Long
        Dim lCalcButtonWidth As Long
        Dim lButtonEdge As Long
        Dim lButtonGap As Long
        Dim lPromptLeft As Long
        Dim lPromptRightGap As Long
        Dim strTemp As String
        Dim bSpaceOutButtons As Boolean
        Dim lFormWidthFromButtons As Long
        Dim lButtonSpacer As Long
        Dim bFormWider As Boolean
        Dim lVerticalGap As Long
        Dim lButtonHeight As Long
        Dim lIconTop As Long
        Dim lIconWidth As Long
        Dim lIconHeight As Long
        Dim lIconGap As Long
        Dim lFormHeightCorrection As Long
        Dim lBreakPos As Long

10      On Error GoTo ERROROUT

20      Font.Name = "Tahoma"
30      Font.Size = m_PromptButtonFontSize

40      lButtonGap = m_MinButtonGap
50      lButtonEdge = m_ButtonEdge
60      lButtonFontColour = GetContrastingFontx(m_ColourButton)
70      lActiveButtonFontColour = GetContrastingFontx(m_ColourActiveButton)
80      bSpaceOutButtons = m_SpaceButtons
90      lPromptRightGap = 200   'this can't be customised (yet?)
100     lVerticalGap = 120
110     lButtonHeight = 375
120     lIconWidth = 512   'may be worth it to see what this exactly is
130     lIconHeight = 512
140     lIconGap = 100

150     If m_CenterPrompt Then
160       txtPrompt(0).Alignment = 2
170     Else
180       txtPrompt(0).Alignment = 0
190     End If

200     If m_Icon = 0 Then
210       lPromptLeft = lPromptRightGap
220       txtPrompt(0).Left = lPromptLeft
230     Else
240       lPromptLeft = lPromptRightGap + lIconWidth + lIconGap
250       txtPrompt(0).Left = lPromptLeft
260     End If

        'to avoid button text being cut off
        '----------------------------------
270     If lButtonEdge < 150 Then
280       lButtonEdge = 150
290     End If

300     If m_ShowTitleIcon = True Then
310       BorderStyle = 3
320     Else
330       BorderStyle = 4
340     End If

350     Caption = m_Title

360     ReDim arrButtonResults(0 To m_ButtonCount - 1) As String

        '360     cmdButton(0).Caption = m_ButtonCaptions(0)

370     For i = 0 To m_ButtonCount - 1
380       cmdButton(i).Caption = m_ButtonCaptions(i)
390       arrButtonResults(i) = Replace(m_ButtonCaptions(i), "&", "")
400     Next

        'strip the shortcut character & off after having done the captions
        '-----------------------------------------------------------------
410     For i = 0 To m_ButtonCount - 1
          '410       If Left$(m_ButtonCaptions(i), 1) = "&" Then
          '420         m_ButtonCaptions(i) = Mid$(m_ButtonCaptions(i), 2)
          '430       End If
420       If m_InputNumeric Then  'PT val
430         If LCase(arrButtonResults(i)) = "cancel" Then
440           cmdButton(i).CausesValidation = False
450         End If
460       End If
470     Next

480     lMaxButtonWidth = TextWidth(arrButtonResults(0)) + lButtonEdge

490     If m_ButtonCount > 1 Then
500       For i = 1 To m_ButtonCount - 1
510         If TextWidth(arrButtonResults(i)) + lButtonEdge > lMaxButtonWidth Then
520           lMaxButtonWidth = TextWidth(arrButtonResults(i)) + lButtonEdge
530         End If
540       Next
550     End If

560     If lMaxButtonWidth > m_MinButtonWidth Then
570       lCalcButtonWidth = lMaxButtonWidth
580     Else
590       lCalcButtonWidth = m_MinButtonWidth
600     End If

        'lCalcButtonWidth now includes the ButtonEdge
        '--------------------------------------------

        'to avoid setting the focus to a non-existing button
        '---------------------------------------------------
610     If m_Default > m_ButtonCount Then
620       m_Default = 1
630     End If

640     If m_PromptBorder Then
650       txtPrompt(0).BorderStyle = 1
660     Else
670       txtPrompt(0).BorderStyle = 0
680     End If

690     txtPrompt(0).Text = m_Prompt

700     If m_ColourForm > -1 Then
710       BackColor = m_ColourForm Or &H2000000
720     End If

730     If m_ColourLabel > -1 Then
740       txtPrompt(0).BackColor = m_ColourLabel Or &H2000000
750     Else
          'this is needed for if no colour is specified at all
760       txtPrompt(0).BackColor = BackColor
770     End If

780     If m_ColourLabelFont > -1 Then
790       txtPrompt(0).ForeColor = m_ColourLabelFont
800     End If

810     If m_ColourButton > -1 Then
820       cmdButton(0).BackColor = m_ColourButton Or &H2000000
830       cmdButton(0).ForeColor = lButtonFontColour
840     End If

850     cmdButton(0).Width = lCalcButtonWidth
860     cmdButton(0).Height = lButtonHeight
870     lFormWidth = lCalcButtonWidth + lPromptLeft + lPromptRightGap

880     If m_ButtonCount > 1 Then
890       For i = 1 To m_ButtonCount - 1
900         lFormWidth = _
            lCalcButtonWidth * (i + 1) + lButtonGap * i + lPromptLeft + lPromptRightGap
910         cmdButton(i).Width = lCalcButtonWidth
920         cmdButton(i).Height = lButtonHeight

            'PT
            'set tabindex's, makes the arrow keys more logical ?
            'has effect here of making moving textbox tabindex's after the buttons
            'In Form_Activate the textbox(1) tabindex might be set to front
930         cmdButton(i).TabIndex = i  'PT
940         If m_ColourButton > -1 Then
950           cmdButton(i).BackColor = m_ColourButton Or &H2000000
960           cmdButton(i).ForeColor = lButtonFontColour
970         End If
980         cmdButton(i).Visible = True
990       Next
1000    End If

        'see if the Close button should be disabled
        '------------------------------------------
1010    If m_ButtonCount > 1 Then
1020      For i = 0 To m_ButtonCount - 1
1030        If LCase(arrButtonResults(i)) = "cancel" Then
1040          bCanEscape = True
1050          Exit For
1060        End If
1070      Next
1080    Else
1090      bCanEscape = True
1100    End If

1110    If bCanEscape = False Then
1120      DisableCloseButton hwnd
1130    End If

        'this is the form width as calculated from the buttons only
        'this will be the buttons, gaps and left and right space, but
        'not the form edges (= lExtraWidth)
        '----------------------------------------------------------
1140    lFormWidthFromButtons = lFormWidth

1150    txtPrompt(0).Font.Size = m_PromptButtonFontSize

1160    For i = 0 To m_ButtonCount - 1
1170      cmdButton(i).Font.Size = m_PromptButtonFontSize
1180    Next

1190    strTemp = m_Prompt

1200    If m_PromptBorder Then
1210      lPromptHeightCorrection = 72
1220      lPromptWidthCorrection = 300
1230      txtPrompt(0).Top = lVerticalGap + 36
1240      lIconTop = lVerticalGap + 36
1250    Else
1260      txtPrompt(0).Top = lVerticalGap
1270      lPromptHeightCorrection = 0
1280      lPromptWidthCorrection = 200
1290    End If

1300    FontBold = True
1310    lExtraHeight = Height - ScaleHeight
1320    lExtraWidth = Width - ScaleWidth
1330    lIconsCorrection = lExtraHeight * 2
1340    lCaptionWidth = TextWidth(Caption) + lIconsCorrection
1350    Font.Bold = False
1360    btTabWidthInSpaces = TextWidth(vbTab) \ TextWidth(Chr(32))
        'without this replace the width of tabs is under-calculated
        '----------------------------------------------------------
1370    lPromptWidth = TextWidth(Replace(strTemp, _
                                         vbTab, _
                                         String(btTabWidthInSpaces, Chr(32)), _
                                         1, -1, _
                                         vbBinaryCompare)) + _
                                         lPromptWidthCorrection

1380    If Len(m_InputDefault) > 0 Then
1390      lInputWidth = TextWidth(m_InputDefault) + 300
1400      If lInputWidth > lPromptWidth Then
1410        lPromptWidth = lInputWidth
1420      End If
1430    End If

1440    lPromptHeight = TextHeight(strTemp) + lPromptHeightCorrection

1450    lFormWidth = lFormWidth + lExtraWidth

        'correct form for prompt width, lFormWidth was set by buttons
        '------------------------------------------------------------
1460    If lPromptWidth + lPromptLeft + lPromptRightGap > lFormWidth Then
1470      lFormWidth = lPromptWidth + lPromptLeft + lPromptRightGap
1480      bFormWider = True
1490    End If

        'correct form for title caption width
        '------------------------------------
1500    If lCaptionWidth > lFormWidth Then
1510      If lCaptionWidth > Screen.Width - 2000 Then
1520        lFormWidth = Screen.Width - 2000
1530      Else
1540        lFormWidth = lCaptionWidth
1550      End If
1560      bFormWider = True
1570    End If

1580    txtPrompt(0).Height = lPromptHeight
1590    txtPrompt(0).Width = (lFormWidth - (lPromptLeft + lPromptRightGap)) - lExtraWidth

1600    If m_DoInput Then
1610      txtPrompt(1).Visible = True
1620      txtPrompt(1).Font.Size = m_PromptButtonFontSize
1630      If m_MaskInput Then
1640        txtPrompt(1).PasswordChar = "*"
1650      Else
1660        txtPrompt(1).PasswordChar = ""
1670      End If
1680      lInputHeight = TextHeight("A Test 1134") + 72
1690      For i = 0 To m_ButtonCount - 1
1700        cmdButton(i).Top = txtPrompt(0).Top + txtPrompt(0).Height + _
                               lVerticalGap * 2 + lInputHeight + 36
1710      Next
1720    Else
1730      For i = 0 To m_ButtonCount - 1
1740        cmdButton(i).Top = txtPrompt(0).Top + txtPrompt(0).Height + _
                               lVerticalGap
1750      Next
1760    End If

1770    lFormHeight = lExtraHeight + cmdButton(0).Top + lButtonHeight + lVerticalGap

        'prevent the form going off the screen at the bottom
        '---------------------------------------------------
1780    If lFormHeight > Screen.Height Then
1790      lFormHeightCorrection = (lFormHeight - Screen.Height) + lExtraHeight
1800      txtPrompt(0).Height = txtPrompt(0).Height - lFormHeightCorrection
1810      For i = 0 To m_ButtonCount - 1
1820        cmdButton(i).Top = cmdButton(i).Top - lFormHeightCorrection
1830      Next
1840      lFormHeight = lFormHeight - lFormHeightCorrection
1850    End If

1860    If m_DoInput Then
1870      With txtPrompt(1)
1880        .BackColor = vbWhite
1890        .BorderStyle = 1
1900        .ForeColor = &H80000008
1910        .Alignment = 0
1920        .Text = m_InputDefault
1930        .Height = lInputHeight
1940        .Width = txtPrompt(0).Width
1950        .Left = txtPrompt(0).Left
1960        .Top = txtPrompt(0).Top + txtPrompt(0).Height + lVerticalGap
1970        .Locked = False
1980        .SelStart = 0
1990        .SelLength = Len(m_InputDefault)
2000      End With
2010    End If

2020    If txtPrompt(0).Height > lIconHeight Then
2030      lIconTop = (txtPrompt(0).Top + txtPrompt(0).Height / 2) - lIconHeight / 2
2040    End If

2050    Height = lFormHeight
2060    Width = lFormWidth

2070    lMidForm = txtPrompt(0).Left + txtPrompt(0).Width / 2

2080    If m_DoInput Then
          'to put the buttons to the right as a regular inputbox
          '-----------------------------------------------------
2090      cmdButton(1).Left = (txtPrompt(1).Left + txtPrompt(1).Width) - lCalcButtonWidth
2100      cmdButton(0).Left = cmdButton(1).Left - (lButtonGap + lCalcButtonWidth)
2110    Else
2120      If bSpaceOutButtons And bFormWider Then
            'spacing the buttons out
            '-----------------------
2130        Select Case m_ButtonCount
              Case 1
2140            cmdButton(0).Left = lMidForm - lCalcButtonWidth / 2
2150          Case 2
2160            lButtonSpacer = (txtPrompt(0).Width - lCalcButtonWidth * 2) / 3
2170            cmdButton(0).Left = lPromptLeft + lButtonSpacer
2180            cmdButton(1).Left = cmdButton(0).Left + lCalcButtonWidth + lButtonSpacer
2190          Case 3
2200            lButtonSpacer = (txtPrompt(0).Width - lCalcButtonWidth * 3) / 4
2210            cmdButton(0).Left = lPromptLeft + lButtonSpacer
2220            cmdButton(1).Left = lMidForm - lCalcButtonWidth / 2
2230            cmdButton(2).Left = cmdButton(1).Left + lCalcButtonWidth + lButtonSpacer
2240          Case 4
2250            lButtonSpacer = (txtPrompt(0).Width - lCalcButtonWidth * 4) / 5
2260            cmdButton(0).Left = lPromptLeft + lButtonSpacer
2270            cmdButton(1).Left = cmdButton(0).Left + lCalcButtonWidth + lButtonSpacer
2280            cmdButton(2).Left = cmdButton(1).Left + lCalcButtonWidth + lButtonSpacer
2290            cmdButton(3).Left = cmdButton(2).Left + lCalcButtonWidth + lButtonSpacer
2300        End Select
2310      Else
            'fixed button gap
            '----------------
2320        Select Case m_ButtonCount
              Case 1
2330            cmdButton(0).Left = lMidForm - lCalcButtonWidth / 2
2340          Case 2
2350            cmdButton(0).Left = lMidForm - (lButtonGap / 2 + lCalcButtonWidth)
2360            cmdButton(1).Left = lMidForm + lButtonGap / 2
2370          Case 3
2380            cmdButton(0).Left = lMidForm - (lCalcButtonWidth / 2 + lButtonGap + lCalcButtonWidth)
2390            cmdButton(1).Left = lMidForm - lCalcButtonWidth / 2
2400            cmdButton(2).Left = lMidForm + lCalcButtonWidth / 2 + lButtonGap
2410          Case 4
2420            cmdButton(0).Left = lMidForm - (lButtonGap / 2 + lCalcButtonWidth * 2 + lButtonGap)
2430            cmdButton(1).Left = lMidForm - (lButtonGap / 2 + lCalcButtonWidth)
2440            cmdButton(2).Left = lMidForm + lButtonGap / 2
2450            cmdButton(3).Left = lMidForm + lButtonGap / 2 + lCalcButtonWidth + lButtonGap
2460        End Select
2470      End If
2480    End If

2490    Left = (Screen.Width - Width) / 2
2500    Top = (Screen.Height - Height) / 2

2510    If m_DoInput And Len(m_InputDefault) > 0 Then
2520      On Error Resume Next
2530      With txtPrompt(1)
2540        .SetFocus
2550        .SelStart = 0
2560        .SelLength = Len(m_InputDefault)
2570      End With
2580    End If

2590    If m_DoInput Then
2600      If Len(m_InputDefault) = 0 Then
2610        On Error Resume Next
2620        txtPrompt(1).SetFocus
2630      End If
2640    End If

2650    bReadyForInput = True

2660    Exit Sub
ERROROUT:

2670    lBreakPos = InStr(1, m_Prompt, vbCrLf, vbBinaryCompare)
2680    If lBreakPos = 0 Then
2690      lBreakPos = Len(m_Prompt)
2700    End If

2710    MsgBox "Form Activate" & vbCrLf & _
               "Line: " & Erl & vbCrLf & _
               Err.Description & vbCrLf & _
               "Error number: " & Err.Number & vbCrLf & vbCrLf & _
               "First line of message prompt:" & vbCrLf & _
               Left$(m_Prompt, lBreakPos)
2720    Stop
2730    Resume
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode = vbFormControlMenu Then

    Dim i As Integer

    If m_ButtonCount = 1 Then
      s_RetVal = arrButtonResults(0)
      Exit Sub
    End If

    For i = 0 To m_ButtonCount - 1
      If LCase$(arrButtonResults(i)) = "cancel" Then
        s_RetVal = arrButtonResults(i)
        Me.Hide
        Exit Sub
      End If
    Next

    'this shouldn't happen as if there is more than one button
    'the close button will be de-activated, unless there is a Cancel button,
    'in which case the caption of that button will be returned
    '-----------------------------------------------------------------------
    s_RetVal = ""

  End If

End Sub
