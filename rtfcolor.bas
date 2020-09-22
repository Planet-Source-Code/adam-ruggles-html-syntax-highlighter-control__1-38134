Attribute VB_Name = "mNew"
Option Explicit
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63
Public varColorText, varColorTag, varColorProp, varColorPropVal, varColorComment, varColorDB As OLE_COLOR
Public regex1 As RegExp, regex2 As RegExp, regex3 As RegExp, regex4 As RegExp ' regex for highlighting
Public matches As MatchCollection
Public match As match
Dim cAppendStr As cAppendString
Public rtfheader As String

Function fcnGetRTFColor(ByVal Color As Variant) As String
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
  
  fcnGetRTFColor = "\red" & CStr(lngRed) & "\green" & CStr(lngGreen) & "\blue" & CStr(lngBlue) & ";"
End Function

Public Function colorhtml(htmltext As String) As String
Set cAppendStr = New cAppendString

cAppendStr.Append rtfheader
'Escape Meta RTF chars using Regular expression
htmltext = regex1.Replace(htmltext, "\$1")
htmltext = regex2.Replace(htmltext, "\par \r")
' color text using regular expressions
    Set matches = regex3.Execute(htmltext)    ' Execute search.
    For Each match In matches     ' Iterate Matches collection.
    'If match <> "" Then
      'With DB
      'cAppendStr.Append "\plain\f2\fs20\cf0 " & match.SubMatches(0) & "\plain\f2\fs20\cf4 " & match.SubMatches(1) & "\plain\f2\fs20\cf5" & match.SubMatches(3) & "\plain\f2\fs20\cf1"
      'cAppendStr.Append regex4.Replace(match.SubMatches(6), "\plain\f2\fs20\cf2 $1\plain\f2\fs20\cf3 $2\plain\f2\fs20\cf1 ")
      'Without DB
      cAppendStr.Append "\plain\f2\fs20\cf0 " & match.SubMatches(0) & "\plain\f2\fs20\cf4 " & match.SubMatches(1) & "\plain\f2\fs20\cf1"
      cAppendStr.Append regex4.Replace(match.SubMatches(3), "\plain\f2\fs20\cf2 $1\plain\f2\fs20\cf3 $2\plain\f2\fs20\cf1 ")
    
    'End If
Next
'add footer
cAppendStr.Append "}"
colorhtml = cAppendStr.Value
cAppendStr.Clear
End Function
