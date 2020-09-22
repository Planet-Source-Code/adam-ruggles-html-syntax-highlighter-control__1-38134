VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl HTMLControl 
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9315
   ScaleHeight     =   7500
   ScaleWidth      =   9315
   Begin RichTextLib.RichTextBox RichTxtBox 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11245
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"HTMLControl.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7245
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   10769
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "Redo"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "HTMLControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" _
        (ByVal hwndLock As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
        
Private Const WM_USER = &H400
Private Const EM_HIDESELECTION = WM_USER + 63
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const WM_COPY = &H301
Private Const WM_CUT = &H300
Private Const WM_CLEAR = &H303
Private Const WM_PASTE = &H302
Private Const EM_SETTARGETDEVICE = WM_USER + 72
'flag to indicate whether actions should be trapped
Private TrapUndo As Boolean
Private UndoStack As New Collection 'collection of undo elements
Private RedoStack As New Collection 'collection of redo elements
'Colors assigned for highlighting
Private varColorText As OLE_COLOR
Private varColorTag As OLE_COLOR
Private varColorProp As OLE_COLOR
Private varColorPropVal As OLE_COLOR
Private varColorComment As OLE_COLOR
'For regular expression calcs
Private regexpRmMeta As RegExp
Private regexpRmNl As RegExp
Private regexpProp As RegExp
Private regexpTags As RegExp
Private regexpComments As RegExp
'To store matches in ColorHtml
Private cMatches As MatchCollection
Private Matches As Match

Private cAppendStr As cAppendString
Private rtfheader As String

Private WasInComment As Boolean

Const def_varColorText = vbBlack
Const def_varColorTag = &HC00000
Const def_varColorProp = &HC000C0
Const def_varColorPropVal = &HC000&
Const def_varColorComment = vbRed

Private Type MatchArray
  FirstIndex As Long
  Length As Long
  Value As String
End Type

Option Explicit
Public Property Let SetStatusBar(ByVal strNewText As String)
  sbStatusBar.Panels(1).Text = strNewText
End Property

Public Property Get SelText() As String
  SelText = RichTxtBox.SelText
End Property

Public Sub CopyText()
  mnuEditCopy_Click
End Sub

Public Sub CutText()
  mnuEditCut_Click
End Sub

Public Sub PasteText()
  mnuEditPaste_Click
End Sub

Public Property Let SelText(ByVal strNewValue As String)
  Screen.MousePointer = vbHourglass
  RichTxtBox.SelRTF = ColorHtml(strNewValue)
  Screen.MousePointer = vbNormal
End Property

Public Property Let SelStart(ByVal newSelStart As Long)
  RichTxtBox.SelStart = newSelStart
End Property

Public Property Let SelLength(ByVal newSelLength As Long)
  RichTxtBox.SelLength = newSelLength
End Property

Public Property Get SelStart() As Long
  SelStart = RichTxtBox.SelStart
End Property

Public Property Get SelLength() As Long
  SelLength = RichTxtBox.SelLength
End Property

Public Property Let SetWidth(newWidth As Long)
  UserControl.Width = newWidth
End Property
Public Property Let SetHeight(newHeight As Long)
  UserControl.Height = newHeight
End Property

Public Property Get Text() As String
   Text = RichTxtBox.Text
End Property

Public Property Get SelRTF() As String
   SelRTF = RichTxtBox.SelRTF
End Property

Public Property Get TextRTF() As String
   TextRTF = RichTxtBox.TextRTF
End Property

Public Property Let TextRTF(newTxtRTF As String)
   RichTxtBox.TextRTF = newTxtRTF
End Property

Public Property Let Text(ByVal strNewValue As String)
   RichTxtBox.Text = strNewValue
End Property
' ******    IMPORTANT INFO ABOUT FONTS    ******
'  -- After Setting new font information run --
'  -- SetNewFontInfo to rebuild and refresh  --
'  -- the text box.                          --
Public Property Let FontName(ByVal strNewValue As String)
  RichTxtBox.Font.Name = strNewValue
End Property

Public Property Get FontName() As String
  FontName = RichTxtBox.Font.Name
End Property

Public Property Let FontSize(ByVal strNewValue As Long)
  RichTxtBox.Font.Size = strNewValue
End Property

Public Property Get FontSize() As Long
  FontSize = RichTxtBox.Font.Size
End Property

Public Property Let FontBold(ByVal strNewValue As Boolean)
  RichTxtBox.Font.Bold = strNewValue
End Property

Public Property Get FontBold() As Boolean
  FontBold = RichTxtBox.Font.Bold
End Property

Public Property Let FontItalic(ByVal strNewValue As Boolean)
  RichTxtBox.Font.Italic = strNewValue
End Property

Public Property Get FontItalic() As Boolean
  FontItalic = RichTxtBox.Font.Italic
End Property

Public Sub SetNewFontInfo()
  BuildRTFHeader
End Sub

Private Sub mnuEditCopy_Click()
  EditFunction WM_COPY
  RichTxtBox.SetFocus
End Sub

Private Sub mnuEditCut_Click()
  EditFunction WM_CUT
  RichTxtBox.SetFocus
End Sub
Private Sub EditFunction(Action As Integer)
  Call SendMessage(RichTxtBox.hwnd, Action, 0, 0&)
  If Action <> WM_COPY Then RichTxtBox.SelText = ""
End Sub

Private Sub mnuEditPaste_Click()
  Screen.MousePointer = vbHourglass
  RichTxtBox.SelRTF = ColorHtml(Clipboard.GetText(vbCFText))
  Screen.MousePointer = vbNormal
End Sub

Private Sub mnuEditRedo_Click()
  Redo
End Sub

Private Sub mnuEditSelectAll_Click()
  RichTxtBox.SelStart = 0
  RichTxtBox.SelLength = Len(RichTxtBox.Text)
  RichTxtBox.SetFocus
End Sub

Private Sub mnuEditUndo_Click()
  Undo
End Sub

Private Sub RichTxtBox_Change()
  If Not TrapUndo Then Exit Sub 'because trapping is disabled
    
  Dim newElement As New UndoElement   'create new undo element
  Dim c%, l&

  'remove all redo items because of the change
  For c% = 1 To RedoStack.Count
    RedoStack.Remove 1
  Next c%

  'set the values of the new element
  newElement.SelStart = RichTxtBox.SelStart
  newElement.TextLen = Len(RichTxtBox.Text)
  newElement.Text = RichTxtBox.Text

  'add it to the undo stack
  UndoStack.Add Item:=newElement
    
  EnableControls
End Sub

Private Sub EnableControls()
  UserControl.mnuEditUndo.Enabled = UndoStack.Count > 1
  UserControl.mnuEditRedo.Enabled = RedoStack.Count > 0
    
  UserControl.mnuEditUndo.Enabled = UserControl.mnuEditUndo.Enabled
  UserControl.mnuEditRedo.Enabled = UserControl.mnuEditRedo.Enabled
    
  RichTxtBox_SelChange
End Sub

Private Sub RichTxtBox_GotFocus()
  On Error Resume Next
  Dim Control As Control
  For Each Control In Controls
    Control.TabStop = False
  Next Control
End Sub

Private Sub RichTxtBox_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim oldSP As Long
  If Shift = vbCtrlMask Then
    Select Case KeyCode
    Case vbKeyV
      ' User pressed Ctrl+V  - Paste
      Screen.MousePointer = vbHourglass
      RichTxtBox.SelRTF = ColorHtml(Clipboard.GetText(vbCFText))
      Screen.MousePointer = vbNormal
      Shift = 0
      KeyCode = 0
    Case vbKeyA
      ' User pressed Ctrl+A   - Select All
      RichTxtBox.SelStart = 0
      RichTxtBox.SelLength = Len(RichTxtBox.Text)
      RichTxtBox.SetFocus
      Shift = 0
      KeyCode = 0
    Case vbKeyC
      ' User pressed Ctrl+C  - Copy
      EditFunction WM_COPY
      RichTxtBox.SetFocus
      Shift = 0
      KeyCode = 0
    Case vbKeyX
      ' User pressed Ctrl+X  - Cut
      EditFunction WM_CUT
      RichTxtBox.SetFocus
      Shift = 0
      KeyCode = 0
    Case vbKeyZ
      ' User pressed Ctrl+Z  - Undo
      Undo
      Shift = 0
      KeyCode = 0
    End Select
  Else
    Select Case KeyCode
    Case vbKeyTab
      LockWindowUpdate RichTxtBox.hwnd
      RichTxtBox.SelRTF = ColorHtml(vbTab)
      LockWindowUpdate False
      KeyCode = 0
    Case vbKeyF5
      LockWindowUpdate RichTxtBox.hwnd
      oldSP = RichTxtBox.SelStart
      RichTxtBox.TextRTF = ColorHtml(RichTxtBox.Text)
      RichTxtBox.SelStart = oldSP
      LockWindowUpdate False
    End Select
  End If
End Sub

Private Sub RichTxtBox_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  Dim ptrEndComm As Long
  Dim oldSP As Long
  With RichTxtBox
    If InTag = True Then
      Select Case Chr(KeyAscii)
      Case "-"
          ' Check if we are in a comment
          If .SelStart >= 3 Then
            oldSP = .SelStart
            .SelStart = .SelStart - 3
            .SelLength = 3
            If .SelText = "<!-" Then
              .SelColor = varColorComment
            End If
            .SelStart = oldSP
          End If
      Case " "
        'Checks for a property
        If InPropval Then
          .SelColor = varColorPropVal
        Else
          .SelColor = varColorProp
        End If
      Case "="
        'Checks for a Property Value
        .SelText = "="
        .SelColor = varColorPropVal
        KeyAscii = 0
      Case ">"
        'Checks for an End Tag
        If InComment = True Then
          .SelColor = varColorComment
          .SelText = ">"
          KeyAscii = 0
          .SelColor = varColorText
        Else
          .SelColor = varColorTag
          .SelText = ">"
          KeyAscii = 0
          .SelColor = varColorText
        End If
      Case Else
        'Else then make sure you havn't moved
        'into a property value
        If InPropval Then
          .SelColor = varColorPropVal
        End If
        End Select
      If InComment Then .SelColor = varColorComment
    Else
      If InComment Then
        .SelColor = varColorComment
      Else
        If Chr(KeyAscii) = "<" Then
          'Checks for a Start Tag
          .SelColor = varColorTag
        Else
          'Otherwise default text color
          .SelColor = varColorText
        End If
      End If
    End If
  End With
End Sub

Private Sub ConvToRTF(RTB As RichTextBox, OtagSP As Long, OtagLength As Long)
  Dim OrgStart As Long
  Dim OrgLength As Long
  With RTB
    OrgStart = .SelStart
    OrgLength = .SelLength
    .SelStart = OtagSP - 1
    .SelLength = OtagLength + 1
    Screen.MousePointer = vbHourglass
    .SelRTF = ColorHtml(RTB.SelText)
    Screen.MousePointer = vbNormal
    .SelStart = OrgStart
    .SelLength = OrgLength
  End With
End Sub


Private Sub RichTxtBox_MouseDown(Button As Integer, Shift As Integer, _
  x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuEdit
    End If
End Sub

Public Sub GetEditStatus()
   Dim lLine As Long, lCol As Long
   Dim cCol As Long, lChar As Long, i As Long

   lChar = RichTxtBox.SelStart + 1

   lLine = 1 + SendMessageLong(RichTxtBox.hwnd, EM_LINEFROMCHAR, _
           RichTxtBox.SelStart, 0&)

   cCol = SendMessageLong(RichTxtBox.hwnd, EM_LINELENGTH, lChar - 1, 0&)

   i = SendMessageLong(RichTxtBox.hwnd, EM_LINEINDEX, lLine - 1, 0&)
   lCol = lChar - i

   sbStatusBar.Panels(2).Text = lLine
   sbStatusBar.Panels(3).Text = lCol

End Sub

Private Sub RichTxtBox_SelChange()
Dim Ln As Long
    Ln = RichTxtBox.SelLength
    With UserControl
        ' Determine which options are available
        .mnuEditCut.Enabled = Ln
        .mnuEditCopy.Enabled = Ln
        .mnuEditPaste.Enabled = Len(Clipboard.GetText(1))
        .mnuEditSelectAll.Enabled = CBool(Len(RichTxtBox.Text))
    End With
    GetEditStatus
End Sub
Private Sub BuildRegularExpressions()
  ' Build Expression to look for tags- <*>
  Set regexpTags = New RegExp
  regexpTags.Pattern = "(<[^>]*>*)"
  regexpTags.Global = True
  regexpTags.IgnoreCase = False
  
  ' Build Expression to look for comments- <!--*-->
  Set regexpComments = New RegExp
  regexpComments.Pattern = "(<!--[\w\W]*?-->)"
  regexpComments.Global = True
  regexpComments.IgnoreCase = False
  
  ' Build Expression to look for propties and values- property="value"
  Set regexpProp = New RegExp
  regexpProp.Pattern = "(\s\w[\w\d\s:_\-\.]*\s*=\s*)(""[^""]+""|'[^']+'|\d+)"
  regexpProp.IgnoreCase = False
  regexpProp.Global = True
  
  ' Build Expression to remove RTF meta characters
  Set regexpRmMeta = New RegExp
  regexpRmMeta.Pattern = "([{}\\])"
  regexpRmMeta.IgnoreCase = False
  regexpRmMeta.Global = True
  
  ' Build Expression to remove RTF new lines
  Set regexpRmNl = New RegExp
  regexpRmNl.Pattern = "(\r)"
  regexpRmNl.IgnoreCase = False
  regexpRmNl.Global = True
  
  
End Sub

Private Sub UserControl_Initialize()
  varColorText = def_varColorText
  varColorTag = def_varColorTag
  varColorProp = def_varColorProp
  varColorPropVal = def_varColorPropVal
  varColorComment = def_varColorComment
  
  WasInComment = False
  'Turn Word Wrap on
  SendMessageLong RichTxtBox.hwnd, EM_SETTARGETDEVICE, 0, 0
         
  TrapUndo = True     'Enable Undo Trapping
  
  BuildRegularExpressions
  BuildRTFHeader
End Sub

Private Sub BuildRTFHeader()
  Dim holdText As String
  Dim colortbl As String
  Dim rtfcolor(4) As String
  Dim GetHeader As RegExp
  Dim tempStr As String
  holdText = RichTxtBox.Text
  RichTxtBox.Text = ""
  
  ' define fonts/ colors and create rtf header
  rtfcolor(0) = fcnGetRTFColor(varColorText)
  rtfcolor(1) = fcnGetRTFColor(varColorTag)
  rtfcolor(2) = fcnGetRTFColor(varColorProp)
  rtfcolor(3) = fcnGetRTFColor(varColorPropVal)
  rtfcolor(4) = fcnGetRTFColor(varColorComment)
  colortbl = Join(rtfcolor, "")

  'Here we'll get the font header info from the rich text box
  Set GetHeader = New RegExp
  GetHeader.Pattern = "\{(.*)\}"
  GetHeader.IgnoreCase = False
  GetHeader.Global = True
  Set cMatches = GetHeader.Execute(RichTxtBox.TextRTF)
  tempStr = GetHeader.Replace(RichTxtBox.TextRTF, "")
  Debug.Print Left(tempStr, Len(tempStr) - 8)
  'Now we add The font header
  rtfheader = cMatches.Item(0) & vbCrLf
  'destroy the objects
  Set cMatches = Nothing
  Set GetHeader = Nothing
  'Now we add our custom color header
  rtfheader = rtfheader & "{\colortbl" & colortbl & "}" & vbCrLf
  rtfheader = rtfheader & Left(tempStr, Len(tempStr) - 8)
  RichTxtBox.TextRTF = ColorHtml(holdText)
End Sub

Private Sub UserControl_InitProperties()
  varColorText = def_varColorText
  varColorTag = def_varColorTag
  varColorProp = def_varColorProp
  varColorPropVal = def_varColorPropVal
  varColorComment = def_varColorComment
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  varColorText = PropBag.ReadProperty("HLTextColor", _
    def_varColorText)
  varColorTag = PropBag.ReadProperty("HLTagColor", _
    def_varColorTag)
  varColorProp = PropBag.ReadProperty("HLPropertyColor", _
    def_varColorProp)
  varColorPropVal = PropBag.ReadProperty("HLPropertyValueColor", _
    def_varColorPropVal)
  varColorComment = PropBag.ReadProperty("HLCommentColor", _
    def_varColorComment)
End Sub

Private Sub UserControl_Resize()
  RichTxtBox.Move 0, 0, ScaleWidth, ScaleHeight - sbStatusBar.Height
End Sub

Private Sub Undo()
Dim chg$, x&
Dim objElement As Object, objElement2 As Object
Dim DeleteFlag As Boolean 'flag as to whether or not
                          'to delete text or append text
  With RichTxtBox
    If UndoStack.Count > 1 And TrapUndo Then 'we can proceed
      TrapUndo = False
      DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < _
                   UndoStack(UndoStack.Count).TextLen
      If DeleteFlag Then  'delete some text
        x& = SendMessage(.hwnd, EM_HIDESELECTION, 1&, 1&)
        Set objElement = UndoStack(UndoStack.Count)
        Set objElement2 = UndoStack(UndoStack.Count - 1)
        .SelStart = objElement.SelStart - _
          (objElement.TextLen - objElement2.TextLen)
        .SelLength = objElement.TextLen - objElement2.TextLen
        .SelText = ""
        x& = SendMessage(.hwnd, EM_HIDESELECTION, 0&, 0&)
      Else 'append something
        Set objElement = UndoStack(UndoStack.Count - 1)
        Set objElement2 = UndoStack(UndoStack.Count)
        chg$ = Change(objElement.Text, objElement2.Text, _
          objElement2.SelStart + 1 + Abs(Len(objElement.Text) - _
          Len(objElement2.Text)))
        .SelStart = objElement2.SelStart
        .SelLength = 0
        .SelText = chg$
        .SelStart = objElement2.SelStart
        If Len(chg$) > 1 And chg$ <> vbCrLf Then
          .SelLength = Len(chg$)
        Else
          .SelStart = .SelStart + Len(chg$)
        End If
      End If
      RedoStack.Add Item:=UndoStack(UndoStack.Count)
      UndoStack.Remove UndoStack.Count
    End If
    EnableControls
    TrapUndo = True
    .SetFocus
  End With
End Sub

Private Sub Redo()
  Dim chg$
  Dim objElement As Object
  Dim DeleteFlag As Boolean 'flag as to whether or not
                            'to delete text or append text
  With RichTxtBox
    If RedoStack.Count > 0 And TrapUndo Then
      TrapUndo = False
      DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(.Text)
      If DeleteFlag Then  'delete last item
        Set objElement = RedoStack(RedoStack.Count)
        .SelStart = objElement.SelStart
        .SelLength = Len(.Text) - objElement.TextLen
        .SelText = ""
      Else 'append something
        Set objElement = RedoStack(RedoStack.Count)
        chg$ = Change(.Text, objElement.Text, objElement.SelStart + 1)
        .SelStart = objElement.SelStart - Len(chg$)
        .SelLength = 0
        .SelText = chg$
        .SelStart = objElement.SelStart - Len(chg$)
        If Len(chg$) > 1 And chg$ <> vbCrLf Then
          .SelLength = Len(chg$)
        Else
          .SelStart = .SelStart + Len(chg$)
        End If
      End If
      UndoStack.Add Item:=objElement
      RedoStack.Remove RedoStack.Count
    End If
    EnableControls
    TrapUndo = True
    .SetFocus
  End With
End Sub

Private Function Change(ByVal lParam1 As String, _
  ByVal lParam2 As String, startSearch As Long) As String
  ' This is for the undo/redo functions
  Dim tempParam$
  Dim d&
  If Len(lParam1) > Len(lParam2) Then 'swap
    tempParam$ = lParam1
    lParam1 = lParam2
    lParam2 = tempParam$
  End If
  d& = Len(lParam2) - Len(lParam1)
  Change = Mid(lParam2, startSearch - d&, d&)
End Function

Private Function fcnGetRTFColor(ByVal Color As Variant) As String
  '***  this function accepts a VB color (long)
  '***  or a HTML color (string) and
  '***  returns a RTF color table def.
  
  Const sHEX = "0123456789ABCDEF"
  Dim lngRed As Long, lngGreen As Long, lngBlue As Long

  If VarType(Color) = vbLong Then
    lngRed = Color Mod 256&
    lngGreen = (Color Mod 65536) \ 256&
    lngBlue = Color \ 65536
  ElseIf VarType(Color) = vbString Then
    '***  the string should be something like this: #D0D5DF
    '***  strip of the right 6 chars
    Color = Right$(Color, 6)
    
    '***  find the position for each char in sHEX. Position is the value
    lngRed = 16& * (InStr(1, sHEX, Mid$(Color, 1, 1), vbTextCompare) - 1) + _
              1& * (InStr(1, sHEX, Mid$(Color, 2, 1), vbTextCompare) - 1)
    lngGreen = 16& * (InStr(1, sHEX, Mid$(Color, 3, 1), vbTextCompare) - 1) + _
                1& * (InStr(1, sHEX, Mid$(Color, 4, 1), vbTextCompare) - 1)
    lngBlue = 16& * (InStr(1, sHEX, Mid$(Color, 5, 1), vbTextCompare) - 1) + _
               1& * (InStr(1, sHEX, Mid$(Color, 6, 1), vbTextCompare) - 1)
  Else
    '***  this function accepts a VB color (long)
    '***  or a HTML color (string) only.
    Stop
  End If
  fcnGetRTFColor = "\red" & CStr(lngRed) & _
                   "\green" & CStr(lngGreen) & _
                   "\blue" & CStr(lngBlue) & ";"
End Function
Private Function Insert(originalString As String, _
  insertString As String, Start As Long, LengthtoRemove As Long)
  'Inserts one string within another and removing as much as needed from
  'the original string
  Dim toEnd As Long
  Dim toRight As Long
  toEnd = Start + LengthtoRemove
  toRight = Len(originalString) - toEnd
  Insert = Left(originalString, Start) & insertString & _
    Right(originalString, toRight)
End Function

Private Function ColorHtml(htmlText As String) As String
  Dim holdMArray() As String
  Dim offset As Long
  Dim index As Long
  Dim tempStr As String
  ReDim holdMArray(0)
  Set cAppendStr = New cAppendString
  'Adds the header info
  cAppendStr.Append rtfheader
  'Removes the Meta info
  htmlText = regexpRmMeta.Replace(htmlText, "\$1")
  'Removes the New Line info
  htmlText = regexpRmNl.Replace(htmlText, "\par \r")
  'Stores the unmodified comments
  index = 1
  Set cMatches = regexpComments.Execute(htmlText)
  For Each Matches In cMatches
      ReDim Preserve holdMArray(index)
      holdMArray(index) = Matches.Value
      index = index + 1
  Next
  'Adds the RTF color tags for properties and values
  htmlText = regexpProp.Replace(htmlText, "\cf2 $1\cf3 $2\cf1 ")
  'Adds the RTF color tags for the tags
  htmlText = regexpTags.Replace(htmlText, "\cf1 $1\cf0 ")
  'Overwrites the highlighted comments with the unhighlighted ones
  Set cMatches = regexpComments.Execute(htmlText)
  index = 1
  For Each Matches In cMatches
    tempStr = "\cf4 " & holdMArray(index) & "\cf0 "
    htmlText = Insert(htmlText, tempStr, _
      Matches.FirstIndex + offset, Matches.Length)
    index = index + 1
    offset = offset + (Len(tempStr) - Len(Matches.Value))
  Next
  'add modified document
  cAppendStr.Append htmlText
  'add footer
  cAppendStr.Append "}"
  ColorHtml = cAppendStr.Value
  'Clean up
  cAppendStr.Clear
  Set Matches = Nothing
  Set cMatches = Nothing
End Function

Private Function InTag() As Boolean
  'Returns True if you are in a tag
  With RichTxtBox
    If .SelStart > 0 Then
      If InStrRev(.Text, "<", .SelStart, vbBinaryCompare) > InStrRev(.Text, _
        ">", .SelStart, vbBinaryCompare) Then InTag = True
    End If
  End With
End Function
Private Function InComment() As Boolean
  'Returns True if you are in a comment
  With RichTxtBox
    If .SelStart > 0 Then
      If InStrRev(.Text, "<!--", .SelStart, vbBinaryCompare) > _
         InStrRev(.Text, "-->", .SelStart, vbBinaryCompare) _
         Then InComment = True
    End If
  End With
End Function
Private Function InPropval() As Boolean
  'Returns True if you are in a property value
  Dim x, y As Long
  InPropval = False
  With RichTxtBox
    x = InStrRev(.Text, """", .SelStart, vbBinaryCompare)
    y = InStrRev(.Text, "=", .SelStart, vbBinaryCompare)
    If x > y Then
      If InStrRev(.Text, """", x - 1, vbBinaryCompare) < _
        InStrRev(.Text, "=", x - 1, vbBinaryCompare) Then InPropval = True
    ElseIf y = .SelStart Then
      InPropval = True
    End If
  End With
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("HLTextColor", _
    varColorText, def_varColorText)
  Call PropBag.WriteProperty("HLTagColor", _
    varColorTag, def_varColorTag)
  Call PropBag.WriteProperty("HLPropertyColor", _
    varColorProp, def_varColorProp)
  Call PropBag.WriteProperty("HLPropertyValueColor", _
    varColorPropVal, def_varColorPropVal)
  Call PropBag.WriteProperty("HLCommentColor", _
    varColorComment, def_varColorComment)
End Sub
