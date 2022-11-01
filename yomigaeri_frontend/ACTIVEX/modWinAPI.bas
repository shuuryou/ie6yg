Attribute VB_Name = "modWinAPI"
Option Explicit

' Everthing already included in oleexp.tlb is not explicitly mentioned here again.

Public Declare Function AccessibleObjectFromWindow Lib "oleacc.dll" (ByVal hWnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long
Public Declare Function DestroyIcon Lib "USER32.DLL" (ByVal hIcon As Long) As Long
Public Declare Function GetClientRect Lib "USER32.DLL" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetParent Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long
Public Declare Function FindWindowEx Lib "USER32.DLL" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function InflateRect Lib "USER32.DLL" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function LoadImage Lib "USER32.DLL" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowPos Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const IIDSTR_IHTMLElement As String = "{3050f1ff-98b5-11cf-bb82-00aa00bdce0b}"
Public Const IIDSTR_IWebBrowser2 As String = "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"

Public Const OBJID_CLIENT As Long = -4

Public Const TB_BUTTONCOUNT As Long = (WM_USER + &H18)
Public Const TB_GETBUTTON As Long = (WM_USER + &H17)
Public Const TB_ENABLEBUTTON  As Long = (WM_USER + &H1)

' Shift messages this amount to filter in WndProc
Public Const SB_IE6YG_SHIFT As Long = 255

Public Const SB_SETTEXTA As Long = WM_USER + 1
Public Const SB_SETPARTS As Long = WM_USER + 4
Public Const SB_GETPARTS As Long = WM_USER + 6
Public Const SB_SETMINHEIGHT As Long = WM_USER + 8
Public Const SB_SIMPLE As Long = WM_USER + 9
Public Const SB_GETRECT As Long = WM_USER + 10
Public Const SB_SETTEXTW As Long = WM_USER + 11
Public Const SB_SETICON As Long = WM_USER + 15

Public Const SBT_NOTABPARSING As Long = &H800

Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_NOACTIVATE As Long = &H10

Public Const IMAGE_ICON = 1

Private Function HIWORD(lValue As Long) As Integer

    If lValue And &H80000000 Then
        HIWORD = (lValue \ 65535) - 1
      Else
        HIWORD = lValue \ 65535
    End If

End Function

Private Function LOWORD(lValue As Long) As Integer

    If lValue And &H8000& Then
        LOWORD = &H8000 Or (lValue And &H7FFF&)
      Else
        LOWORD = lValue And &HFFFF&
    End If

End Function

Public Function MAKEINTRESOURCE(lID As Long)

    MAKEINTRESOURCE = "#" & CStr(MAKELONG(lID, 0))

End Function

Private Function MAKELONG(wLow As Long, wHi As Long) As Long

    If (wHi And &H8000&) Then
        MAKELONG = (((wHi And &H7FFF&) * 65536) Or (wLow And &HFFFF&)) Or &H80000000
      Else
        MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHi))
    End If

End Function

Public Function TrimNull(ByVal text As String) As String

  Dim lngPos As Long

    lngPos = InStr(text, Chr$(0))
    If lngPos > 0 Then
        text = Left$(text, lngPos - 1)
    End If

    TrimNull = text

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-01 16:05)  Decl: 48  Code: 52  Total: 100 Lines
':) CommentOnly: 4 (4%)  Commented: 0 (0%)  Filled: 71 (71%)  Empty: 29 (29%)  Max Logic Depth: 2
