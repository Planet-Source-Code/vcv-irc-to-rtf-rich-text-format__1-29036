<div align="center">

## IRC TO RTF \(Rich Text Format\)


</div>

### Description

Converts IRC text to RTF and displays it in a RichText Box. Supports everything.
 
### More Info
 
rtf As RichTextBox: the richtext box to append the text to

strData As String: the string to parse

To call it, let's say you had RichTextBox1, you'd use...

PutText RichTextBox1, strIRCText

where strIRCText is the text to parse.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[vcv](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vcv.md)
**Level**          |Intermediate
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vcv-irc-to-rtf-rich-text-format__1-29036/archive/master.zip)

### API Declarations

```
Public Declare Function GetTextCharset Lib "gdi32" (ByVal hdc As Long) As Long
'* ANSI Formatting character values
Global Const Cancel = 15
Global Const BOLD = 2
Global Const UNDERLINE = 31
Global Const Color = 3
Global Const REVERSE = 22
Global Const ACTION = 1
'* ANSI Formatting characters
Global strBold As String
Global strUnderline As String
Global strColor As String
Global strReverse As String
Global strAction As String
Global Const strFontName = "Courier"
Global Const intFontSize = 8
```


### Source Code

```
Function RAnsiColor(lngColor As Long) As Integer
  Select Case lngColor
    Case RGB(255, 255, 255): RAnsiColor = 0
    Case RGB(0, 0, 0): RAnsiColor = 1
    Case RGB(0, 0, 127): RAnsiColor = 2
    Case RGB(0, 127, 0): RAnsiColor = 3
    Case RGB(255, 0, 0): RAnsiColor = 4
    Case RGB(127, 0, 0): RAnsiColor = 5
    Case RGB(127, 0, 127): RAnsiColor = 6
    Case RGB(255, 127, 0): RAnsiColor = 7
    Case RGB(255, 255, 0): RAnsiColor = 8
    Case RGB(0, 255, 0): RAnsiColor = 9
    Case RGB(0, 148, 144): RAnsiColor = 10
    Case RGB(0, 255, 255): RAnsiColor = 11
    Case RGB(0, 0, 255): RAnsiColor = 12
    Case RGB(255, 0, 255): RAnsiColor = 13
    Case RGB(92, 92, 92): RAnsiColor = 14
    Case RGB(184, 184, 184): RAnsiColor = 15
    Case RGB(0, 0, 0): RAnsiColor = 99
    Case lngForeColor: RAnsiColor = 1
    Case lngBackColor: RAnsiColor = 0
  End Select
End Function
Function ColorTable() As String
  Dim i As Integer, strTable As String
  Dim r As Integer, b As Integer, g As Integer
  strTable = "{\colortbl ;"
  For i = 0 To 15
    Select Case i
      Case 0: r = 255: g = 255: b = 255
      Case 1: r = 0: g = 0: b = 0
      Case 2: r = 0: g = 0: b = 127
      Case 3: r = 0: g = 127: b = 0
      Case 4: r = 255: g = 0: b = 0
      Case 5: r = 127: g = 0: b = 0
      Case 6: r = 127: g = 0: b = 127
      Case 7: r = 255: g = 127: b = 0
      Case 8: r = 255: g = 255: b = 0
      Case 9: r = 0: g = 255: b = 0
      Case 10: r = 0: g = 148: b = 144
      Case 11: r = 0: g = 255: b = 255
      Case 12: r = 0: g = 0: b = 255
      Case 13: r = 255: g = 0: b = 255
      Case 14: r = 92: g = 92: b = 92
      Case 15: r = 184: g = 184: b = 184
      Case Else: r = 0: g = 0: b = 0
    End Select
    strTable = strTable & "\red" & r & "\green" & g & "\blue" & b & ";"
  Next i
  strTable = strTable & "}"
  ColorTable = strTable
End Function
Sub PutText(rtf As RichTextBox, strData As String)
  If strData = "" Then Exit Sub
  '* Variable decs
  Dim i As Long, Length As Integer, strChar As String, strBuffer As String
  Dim clr As Integer, bclr As Integer, dftclr As Integer, strRTFBuff As String
  Dim bbbold As Boolean, bbunderline As Boolean, bbreverse As Boolean, strTmp As String
  Dim lngFC As String, lngBC As String, lngStart As Long, lngLength As Long, strPlaceHolder As String
  '* if not inialized, set font, intialiaze (and also generate color table)
  Dim btCharSet As Long
  Dim strRTF As String
  If rtf.Tag <> "init'd" Then
    rtf.Tag = "init'd"
    strFontName = rtf.Font.Name
    rtf.parent.FontName = strFontName
    btCharSet = GetTextCharset(rtf.parent.hdc)
    strRTF = ""
    strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & strFontName & ";}}" & vbCrLf
    strRTF = strRTF & ColorTable & vbCrLf
    strRTF = strRTF & "\viewkind4\uc1\pard\cf0\fi-" & intIndent & "\li" & intIndent & "\f0\fs" & CInt(intFontSize * 2) & vbCrLf
    strPlaceHolder = "\n"
    For i = 0 To 15
      strRTF = strRTF & "\cf" & i & " " & strPlaceHolder
    Next
    strRTF = strRTF & "}"
    rtf.TextRTF = strRTF
    '* New session for window... call
    '# LogData rtf.Parent.Caption, "blah", strData, True
  Else
    '# LogData rtf.Parent.Caption, "blah", strData, False
  End If
  '* Generate header information to use (font name, size, etc)
  rtf.parent.FontName = strFontName
  btCharSet = GetTextCharset(rtf.parent.hdc)
  strRTF = ""
  strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & strFontName & ";}}" & vbCrLf
  strRTF = strRTF & ColorTable & vbCrLf
  strRTF = strRTF & "\viewkind4\uc1\pard\cf0\fi-" & intIndent & "\li" & intIndent & "\f0\fs" & CInt(intFontSize * 2) & vbCrLf
  '* Reset all codes from previous lines.
  strRTFBuff = "\b0\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1 & "\i0\ulnone "
  dftclr = RAnsiColor(lngForeColor)
  '* Set loop
  Length = Len(strData)
  i = 1
  Do
    strChar = Mid(strData, i, 1)
    '* Check the current character
    Select Case strChar
      Case Chr(Cancel)  'cancel code
        ' Reset all previous formatting
        If Right(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
        lngFC = CStr(RAnsiColor(lngForeColor))
        lngBC = CStr(RAnsiColor(lngBackColor))
        strRTFBuff = strRTFBuff & strBuffer & "\b0\ul0\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1
        strBuffer = ""
        i = i + 1
      Case strBold	' bold
        ' Invert the bold flag, append the buffer of previous text, then bold character
        bbbold = Not bbbold
        If Right(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
        strRTFBuff = strRTFBuff & strBuffer & "\b"
        If bbbold = False Then strRTFBuff = strRTFBuff & "0"
        strBuffer = ""
        i = i + 1
      Case strUnderline	' underline
        ' Invert the underline flag, append the buffer of previous text, then under character
        bbunderline = Not bbunderline
        If Right(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
        strRTFBuff = strRTFBuff & strBuffer & "\ul"
        If bbunderline = False Then strRTFBuff = strRTFBuff & "none"
        strBuffer = ""
        i = i + 1
      Case strReverse
        ' Invert the reverse flag, append the buffer of previous text, then set forecolor and backcolor to inverse
        bbreverse = Not bbreverse
        If Right(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " " ' & strBuffer & "\"
        If bbreverse = False Then
          If Right(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
          strRTFBuff = strRTFBuff & strBuffer & "\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1
        Else
          If Right(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
          strRTFBuff = strRTFBuff & strBuffer & "\cf" & RAnsiColor(lngBackColor) + 1 & "\highlight" & RAnsiColor(lngForeColor) + 1
        End If
        strBuffer = ""
        i = i + 1
      Case strColor
        strTmp = ""
        i = i + 1
        ' check the characters following the color character to find the color we need to set.
        Do Until Not ValidColorCode(strTmp) Or i > Length
          strTmp = strTmp & Mid(strData, i, 1)
          i = i + 1
        Loop
        ' If no color specified (color character alone), reset color, else change forecolor and back color if needed
        strTmp = LeftR(strTmp, 1)
        If strTmp = "" Then
          lngFC = CStr(RAnsiColor(lngForeColor))
          lngBC = CStr(RAnsiColor(lngBackColor))
        Else
          lngFC = LeftOf(strTmp, ",")
          lngFC = CStr(CInt(lngFC))
          If InStr(strTmp, ",") Then
            lngBC = RightOf(strTmp, ",")
            If lngBC <> "" Then lngBC = CStr(CInt(lngBC)) Else lngBC = CStr(RAnsiColor(lngBackColor))
          Else
            lngBC = ""
          End If
        End If
        If lngFC = "" Then lngFC = CStr(lngForeColor)
        lngFC = Int(lngFC) + 1
        If lngBC <> "" Then lngBC = Int(lngBC) + 1
        ' This is where we actually change the color.
        ' We append the current buffer of previous text and then change the color
        If Right(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
        strRTFBuff = strRTFBuff & strBuffer
        strRTFBuff = strRTFBuff & "\cf" & lngFC
        If lngBC <> "" Then strRTFBuff = strRTFBuff & "\highlight" & lngBC
        i = i - 1
        strBuffer = ""
        If i >= Length Then GoTo TheEnd
      Case Else
        ' Not a special code, so just append to the buffer of text
        Select Case strChar
        ' make sure the { } and \ characters are properly displayed, because RTF uses them for special formatting, so we escape them with \
        Case "}", "{", "\"
          strBuffer = strBuffer & "\" & strChar
        Case Else
          strBuffer = strBuffer & strChar
        End Select
        i = i + 1
    End Select
  Loop Until i > Length
TheEnd:
  ' if any data is left of buffer of previous text, then append it to the RTF buffer
  If strBuffer <> "" Then
    strRTFBuff = strRTFBuff & " " & strBuffer
  End If
  ' Set the caret to the end of the text and set the "SelRTF property".
  strRTFBuff = strRTFBuff & vbCrLf
  rtf.selStart = Len(rtf.Text)
  rtf.selLength = 0
  rtf.SelRTF = strRTF & strRTFBuff & vbCrLf & " }" & vbCrLf
  rtf.seltext = vbCrLf
End Sub
Function ValidColorCode(strCode As String) As Boolean
  If strCode = "" Then ValidColorCode = True: Exit Function
  Dim c1 As Integer, c2 As Integer
  If strCode Like "" Or _
    strCode Like "#" Or _
    strCode Like "##" Or _
    strCode Like "#,#" Or _
    strCode Like "##,#" Or _
    strCode Like "#,##" Or _
    strCode Like "#," Or _
    strCode Like "##," Or _
    strCode Like "##,##" Or _
    strCode Like ",#" Or _
    strCode Like ",##" Then
    Dim strCol() As String
    strCol = Split(strCode, ",")
    '
    If UBound(strCol) = -1 Then
      ValidColorCode = True
    ElseIf UBound(strCol) = 0 Then
      If strCol(0) = "" Then strCol(0) = 0
      If Int(strCol(0)) >= 0 And Int(strCol(0)) <= 99 Then
        ValidColorCode = True
        Exit Function
      Else
        ValidColorCode = False
        Exit Function
      End If
    Else
      If strCol(0) = "" Then strCol(0) = lngForeColor
      If strCol(1) = "" Then strCol(1) = 0
      c1 = Int(strCol(0))
      c2 = Int(strCol(1))
      If Int(c2) < 0 Or Int(c2) > 99 Then
        ValidColorCode = False
        Exit Function
      Else
        ValidColorCode = True
        Exit Function
      End If
    End If
    ValidColorCode = True
    Exit Function
  Else
    ValidColorCode = False
    Exit Function
  End If
End Function
Function LeftR(strData As String, intMin As Integer)
  On Error Resume Next
  LeftR = Left(strData, Len(strData) - intMin)
End Function
```

