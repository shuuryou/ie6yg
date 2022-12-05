Attribute VB_Name = "modWinAPI"
Option Explicit

' This file only contains stuff that multiple classes, forms, or modules need to
' function correctly.

' Everthing already included in oleexp.tlb is not explicitly declared here again.

' API calls only used by one class, form, or module are at the top of its file.

Public Declare Function AccessibleObjectFromWindow Lib "OLEACC.DLL" (ByVal hWnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long
Public Declare Function IIDFromString Lib "OLE32.DLL" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long

Public Declare Function FindWindowEx Lib "USER32.DLL" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetParent Lib "USER32.DLL" (ByVal hWnd As Long) As Long

Public Declare Function LoadImage Lib "USER32.DLL" Alias "LoadImageA" (ByVal hinst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Public Declare Function DestroyIcon Lib "USER32.DLL" (ByVal hIcon As Long) As Long

Public Const IMAGE_ICON = 1

Public Declare Function GetClientRect Lib "USER32.DLL" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_SHOWWINDOW = &H40

Private Declare Function GetTempPath Lib "KERNEL32.DLL" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "KERNEL32.DLL" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Const MAX_PATH As Long = 260

' For forms
Public Declare Function SetParent Lib "USER32.DLL" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

' For IEStatusBar; these are here because IEFrame needs them too.
Public Enum StatusBarPanes
    Navigation = 0
    Progress = 1
    Connection = 2
    SSL = 3
    Zone = 4
End Enum
#If False Then ':) Line inserted by Formatter
Private Navigation, Progress, Connection, SSL, Zone ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Public Const STATUSBAR_PANES As Integer = 5

#If False Then
' F***ING VB6 IDE KEEPS CHANGING THE CASE OF THESE RANDOMLY!
' This makes it stop.
Private hWnd
Private Left
Private Right
#End If

Public Function HiWord(lDWord As Long) As Integer

    HiWord = (lDWord And &HFFFF0000) \ &H10000

End Function

Public Function LoWord(lDWord As Long) As Integer

    If lDWord And &H8000& Then
        LoWord = lDWord Or &HFFFF0000
      Else
        LoWord = lDWord And &HFFFF&
    End If

End Function

Public Function MAKEINTRESOURCE(lId As Long)

    MAKEINTRESOURCE = "#" & CStr(MAKELPARAM(lId, 0))

End Function

Public Function MAKELONG(wLow As Long, wHigh As Long) As Long

    MAKELONG = LoWord(wLow) Or (&H10000 * LoWord(wHigh))

End Function

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long

    MAKELPARAM = MAKELONG(wLow, wHigh)

End Function

Public Function TempName() As String

  Dim strTempDir As String
  Dim strBuffer As String
  Dim lngRet As Long

    strBuffer = Space$(MAX_PATH)
    lngRet = GetTempPath(MAX_PATH, strBuffer)

    If lngRet = 0 Then
        Err.Raise Err.LastDllError
    End If

    strTempDir = Left$(strBuffer, lngRet)

    lngRet = GetTempFileName(strTempDir, "iyg", 0, strBuffer)

    If lngRet = 0 Then
        Err.Raise Err.LastDllError
    End If

    TempName = TrimNull(strBuffer)

End Function

Public Function TrimNull(ByVal Text As String) As String

  Dim lngPos As Long

    lngPos = InStr(Text, vbNullChar)
    If lngPos > 0 Then
        Text = Left$(Text, lngPos - 1)
    End If

    TrimNull = Text

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-18 05:21)  Decl: 61  Code: 75  Total: 136 Lines
':) CommentOnly: 9 (6.6%)  Commented: 3 (2.2%)  Filled: 91 (66.9%)  Empty: 45 (33.1%)  Max Logic Depth: 2
