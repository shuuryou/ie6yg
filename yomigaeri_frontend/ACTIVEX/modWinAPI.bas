Attribute VB_Name = "modWinAPI"
Option Explicit

' Everthing already included in oleexp.tlb is not explicitly mentioned here again.

Public Declare Function AccessibleObjectFromWindow Lib "oleacc.dll" (ByVal hwnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long
Public Declare Function IIDFromString Lib "OLE32.DLL" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long

Public Declare Function FindWindowEx Lib "USER32.DLL" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetParent Lib "USER32.DLL" (ByVal hwnd As Long) As Long

Public Declare Function LoadImage Lib "USER32.DLL" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Public Declare Function DestroyIcon Lib "USER32.DLL" (ByVal hIcon As Long) As Long

Public Declare Function InflateRect Lib "USER32.DLL" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "USER32.DLL" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetClientRect Lib "USER32.DLL" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function EnableMenuItem Lib "USER32.DLL" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function DrawMenuBar Lib "USER32.DLL" (ByVal hwnd As Long) As Long

Public Declare Sub Sleep Lib "KERNEL32.DLL" (ByVal dwMilliseconds As Long)

Public Declare Function CommitUrlCacheEntry Lib "WININET.DLL" Alias "CommitUrlCacheEntryA" (ByVal lpszUrlName As String, ByVal lpszLocalFileName As String, ByRef tftExpireTime As FILETIME, ByRef tftLastModifiedTime As FILETIME, ByVal lCacheEntryType As Long, ByVal lpHeaderInfo As Long, ByVal dwHeaderSize As Long, ByVal lpszFileExtension As String, ByVal dwReserved As Long) As Long
Public Declare Function CreateUrlCacheEntry Lib "WININET.DLL" Alias "CreateUrlCacheEntryA" (ByVal lpszUrlName As String, ByVal dwExpectedFileSize As Long, ByVal lpszFileExtension As String, ByVal lpszFileName As String, ByVal dwReserved As Long) As Long

Public Const NORMAL_CACHE_ENTRY As Long = &H1

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type NMTOOLBAR_SHORT
    hdr As NMHDR
    iItem As Long
End Type

Public Const IIDSTR_IHTMLElement As String = "{3050f1ff-98b5-11cf-bb82-00aa00bdce0b}"
Public Const IIDSTR_IWebBrowser2 As String = "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"

Public Const OBJID_CLIENT As Long = -4

' Shift messages this amount to filter in WndProc
Public Const TB_IE6YG_SHIFT As Long = 255

Public Const TB_BUTTONCOUNT As Long = (WM_USER + &H18)
Public Const TB_GETBUTTON As Long = (WM_USER + &H17)
Public Const TB_ENABLEBUTTON  As Long = (WM_USER + &H1)

Public Const TBSTYLE_DROPDOWN As Long = &H8
Public Const TBSTYLE_AUTOSIZE As Long = &H10

' Shift messages this amount to filter in WndProc
Public Const SB_IE6YG_SHIFT As Long = 255

Public Const SB_SETTEXTA As Long = WM_USER + 1
Public Const SB_GETTEXTA = WM_USER + 2
Public Const SB_GETTEXTLENGTHA = WM_USER + 3
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
Public Const SWP_SHOWWINDOW = &H40

Public Const IMAGE_ICON = 1

Public Const PBM_GETPOS As Long = WM_USER + 8
Public Const PBM_SETPOS As Long = WM_USER + 2
Public Const PBM_SETRANGE As Long = WM_USER + 1

Public Const SIZE_RESTORED As Long = 0
Public Const SIZE_MINIMIZED As Long = 1
Public Const SIZE_MAXIMIZED As Long = 2

Public Const MF_BYCOMMAND As Long = &H0
Public Const MF_GRAYED As Long = &H1
Public Const MF_ENABLED As Long = &H0

Public Const TBN_FIRST As Long = -700
Public Const TBN_DROPDOWN As Long = (TBN_FIRST - 10)

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

Public Function MAKEINTRESOURCE(lID As Long)

    MAKEINTRESOURCE = "#" & CStr(MAKELPARAM(lID, 0))

End Function

Public Function MAKELONG(wLow As Long, wHigh As Long) As Long

    MAKELONG = LoWord(wLow) Or (&H10000 * LoWord(wHigh))

End Function

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long

    MAKELPARAM = MAKELONG(wLow, wHigh)

End Function

Public Function TrimNull(ByVal Text As String) As String

  Dim lngPos As Long

    lngPos = InStr(Text, Chr$(0))
    If lngPos > 0 Then
        Text = Left$(Text, lngPos - 1)
    End If

    TrimNull = Text

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-11 04:01)  Decl: 91  Code: 50  Total: 141 Lines
':) CommentOnly: 5 (3.5%)  Commented: 0 (0%)  Filled: 95 (67.4%)  Empty: 46 (32.6%)  Max Logic Depth: 2
