Attribute VB_Name = "modFrontendStyling"
Option Explicit

Private Declare Function SystemParametersInfo Lib "USER32.DLL" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "USER32.DLL" (ByVal nIndex As Long) As Long

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const COLOR_SCROLLBAR As Long = 0
Private Const COLOR_BACKGROUND As Long = 1
Private Const COLOR_ACTIVECAPTION As Long = 2
Private Const COLOR_INACTIVECAPTION As Long = 3
Private Const COLOR_MENU As Long = 4
Private Const COLOR_WINDOW As Long = 5
Private Const COLOR_WINDOWFRAME As Long = 6
Private Const COLOR_MENUTEXT As Long = 7
Private Const COLOR_WINDOWTEXT As Long = 8
Private Const COLOR_CAPTIONTEXT As Long = 9
Private Const COLOR_ACTIVEBORDER As Long = 10
Private Const COLOR_INACTIVEBORDER As Long = 11
Private Const COLOR_APPWORKSPACE As Long = 12
Private Const COLOR_HIGHLIGHT As Long = 13
Private Const COLOR_HIGHLIGHTTEXT As Long = 14
Private Const COLOR_BTNFACE As Long = 15
Private Const COLOR_BTNSHADOW As Long = 16
Private Const COLOR_GRAYTEXT As Long = 17
Private Const COLOR_BTNTEXT As Long = 18
Private Const COLOR_INACTIVECAPTIONTEXT As Long = 19
Private Const COLOR_BTNHIGHLIGHT As Long = 20
Private Const COLOR_2NDACTIVECAPTION As Long = 27
Private Const COLOR_2NDINACTIVECAPTION As Long = 28

Private Const SPI_GETNONCLIENTMETRICS As Integer = 41
Private Const LF_FACESIZE As Integer = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

Public Function GetAcceptLanguage() As String

  Dim strRet As String
  Dim Registry As New CRegistry

    With Registry
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\Microsoft\Internet Explorer\International"
        .ValueType = REG_SZ

        .ValueKey = "AcceptLanguage"

        strRet = .value
    End With

    Set Registry = Nothing

    If strRet = "" Then
        strRet = "<EMPTY>"
    End If

    GetAcceptLanguage = strRet

End Function

Private Function LongToRGBHex(ByVal lLong As Long) As String

  ' by Donald, donald@xbeat.net, 20010910

  Dim bRed As Long
  Dim bGreen As Long
  Dim bBlue As Long
  ' mask out highest byte

    lLong = lLong And &HFFFFFF
    ' extract color bytes
    bRed = lLong And &HFF
    bGreen = (lLong \ &H100) And &HFF
    bBlue = (lLong \ &H10000) And &HFF
    ' reverse bytes
    lLong = bRed * &H10000 + bGreen * &H100 + bBlue
    ' to hex, left-padd zeroes
    ' the string op is the bottleneck of this procedure, and since in real
    ' world most colors have a red-part >= 16, it's a good idea to check if
    ' we really need the string op
    If bRed < &H10 Then
        LongToRGBHex = Right$("00000" & Hex$(lLong), 6)
      Else
        LongToRGBHex = Hex$(lLong)
    End If

End Function

Public Function MakeStyling(ByRef hDC As Long) As String

  Dim lngRet As Long
  Dim sctNCM As NONCLIENTMETRICS
  Dim strFont As String
  Dim strResponse As String

    sctNCM.cbSize = Len(sctNCM)

    lngRet = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, sctNCM, 0)

    If lngRet = 0 Then
        MakeStyling = "ERROR"
        Exit Function
    End If

    strFont = StrConv(sctNCM.lfMessageFont.lfFaceName, vbUnicode)
    strFont = Left$(strFont, InStr(strFont, vbNullChar) - 1)

    ' Str$ in the font size is important to account for locales that
    ' do not use a period as the decimal separator.

    strResponse = strFont & vbTab & _
                  Str$(-(sctNCM.lfMessageFont.lfHeight * (72 / GetDeviceCaps(hDC, LOGPIXELSY)))) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_SCROLLBAR)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_BACKGROUND)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_ACTIVECAPTION)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_INACTIVECAPTION)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_MENU)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_WINDOW)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_WINDOWFRAME)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_MENUTEXT)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_WINDOWTEXT)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_CAPTIONTEXT)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_ACTIVEBORDER)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_INACTIVEBORDER)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_APPWORKSPACE)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_HIGHLIGHT)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_HIGHLIGHTTEXT)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_BTNFACE)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_BTNSHADOW)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_GRAYTEXT)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_BTNTEXT)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_INACTIVECAPTIONTEXT)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_BTNHIGHLIGHT)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_2NDACTIVECAPTION)) & vbTab & _
                  LongToRGBHex(GetSysColor(COLOR_2NDINACTIVECAPTION))

    MakeStyling = strResponse

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-21 08:22)  Decl: 70  Code: 108  Total: 178 Lines
':) CommentOnly: 12 (6.7%)  Commented: 0 (0%)  Filled: 147 (82.6%)  Empty: 31 (17.4%)  Max Logic Depth: 2
