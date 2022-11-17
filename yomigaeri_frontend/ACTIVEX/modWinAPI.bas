Attribute VB_Name = "modWinAPI"
Option Explicit

' Everthing already included in oleexp.tlb is not explicitly mentioned here again.

Public Declare Function AccessibleObjectFromWindow Lib "OLEACC.DLL" (ByVal hwnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long
Public Declare Function IIDFromString Lib "OLE32.DLL" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long

Public Const IIDSTR_IHTMLElement As String = "{3050f1ff-98b5-11cf-bb82-00aa00bdce0b}"
Public Const IIDSTR_IWebBrowser2 As String = "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"

Public Const OBJID_CLIENT As Long = -4

Public Declare Function FindWindowEx Lib "USER32.DLL" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetParent Lib "USER32.DLL" (ByVal hwnd As Long) As Long

Public Declare Function LoadImage Lib "USER32.DLL" Alias "LoadImageA" (ByVal hinst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Public Declare Function DestroyIcon Lib "USER32.DLL" (ByVal hIcon As Long) As Long

Public Declare Function InflateRect Lib "USER32.DLL" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "USER32.DLL" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetClientRect Lib "USER32.DLL" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Declare Function EnableMenuItem Lib "USER32.DLL" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function DrawMenuBar Lib "USER32.DLL" (ByVal hwnd As Long) As Long

Public Declare Function CommitUrlCacheEntry Lib "WININET.DLL" Alias "CommitUrlCacheEntryA" (ByVal lpszUrlName As String, ByVal lpszLocalFileName As String, ByRef tftExpireTime As FILETIME, ByRef tftLastModifiedTime As FILETIME, ByVal lCacheEntryType As Long, ByVal lpHeaderInfo As Long, ByVal dwHeaderSize As Long, ByVal lpszFileExtension As String, ByVal dwReserved As Long) As Long
Public Declare Function CreateUrlCacheEntry Lib "WININET.DLL" Alias "CreateUrlCacheEntryA" (ByVal lpszUrlName As String, ByVal dwExpectedFileSize As Long, ByVal lpszFileExtension As String, ByVal lpszFileName As String, ByVal dwReserved As Long) As Long

Public Const NORMAL_CACHE_ENTRY As Long = &H1

Public Declare Function SetCursor Lib "USER32.DLL" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "USER32.DLL" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Public Declare Function DestroyCursor Lib "USER32.DLL" (ByVal hCursor As Long) As Long

Public Declare Function CreateWindowEx Lib "USER32.DLL" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "USER32.DLL" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "USER32.DLL" (lpPoint As POINT) As Long

Public Declare Function SetParent Lib "USER32.DLL" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Declare Function GetTempPath Lib "KERNEL32.DLL" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "KERNEL32.DLL" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Const MAX_PATH As Long = 260

' For ctlFrontend (cursor update)
Public Const GCL_HCURSOR As Long = (-12)
Public Const IDC_ARROW As Long = &H7F00

' For IEToolTip
Public Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uid As Long
    rc As RECT
    hinst As Long
    lpszText As String
    lParam As Long
End Type

Public Const CW_USEDEFAULT As Long = &H80000000
Public Const TOOLTIPS_CLASSA As String = "tooltips_class32"
Public Const HWND_TOPMOST As Long = -1

Public Const TTF_TRACK As Long = &H20
Public Const TTF_ABSOLUTE As Long = &H80

Public Const TTS_ALWAYSTIP As Long = &H1
Public Const TTS_NOPREFIX As Long = &H2

Public Const TTM_ADDTOOL As Long = (WM_USER + 4)
Public Const TTM_UPDATETIPTEXT As Long = (WM_USER + 12)
Public Const TTM_TRACKACTIVATE As Long = (WM_USER + 17)
Public Const TTM_TRACKPOSITION As Long = (WM_USER + 18)

' For IEToolbar
Public Type NMTOOLBAR_SHORT
    hdr As NMHDR
    iItem As Long
End Type

' Shift messages this amount to filter in WndProc
Public Const TB_IE6YG_SHIFT As Long = 255

Public Const TB_BUTTONCOUNT As Long = (WM_USER + &H18)
Public Const TB_GETBUTTON As Long = (WM_USER + &H17)
Public Const TB_ENABLEBUTTON  As Long = (WM_USER + &H1)

Public Const TBN_FIRST As Long = -700
Public Const TBN_DROPDOWN As Long = (TBN_FIRST - 10)

Public Const TBSTYLE_DROPDOWN As Long = &H8
Public Const TBSTYLE_AUTOSIZE As Long = &H10

' For IEStatusBar
Public Const STATUS_PANES As Integer = 5
Public Const STATUS_PANE_NAVIGATION As Integer = 0
Public Const STATUS_PANE_PROGRESS As Integer = 1
Public Const STATUS_PANE_CONNECTION As Integer = 2
Public Const STATUS_PANE_SSL As Integer = 3
Public Const STATUS_PANE_ZONE As Integer = 4

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
Public Const SB_SETTIPTEXTA As Long = WM_USER + 16

Public Const SBT_NOTABPARSING As Long = &H800

Public Const IMAGE_ICON = 1

Public Const PBM_GETPOS As Long = WM_USER + 8
Public Const PBM_SETPOS As Long = WM_USER + 2
Public Const PBM_SETRANGE As Long = WM_USER + 1

' For IEFrame
Public Declare Function EnableWindow Lib "USER32.DLL" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Public Const SIZE_RESTORED As Long = 0
Public Const SIZE_MINIMIZED As Long = 1
Public Const SIZE_MAXIMIZED As Long = 2

Public Const MF_BYCOMMAND As Long = &H0
Public Const MF_GRAYED As Long = &H1
Public Const MF_ENABLED As Long = &H0

Public Const NM_DBLCLK As Long = -3

Public Type NMMOUSE
    hdr As NMHDR
    dwItemSpec As Long
    dwItemData As Long
    pt As POINT
    dwHitInfo As Long
End Type

' For IEAddressBar
Public Type COMBOBOXEXITEM
    mask As Long ' CBEIF..
    iItem As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    iOverlay As Long
    iIndent As Long
    lParam As Long
End Type

Public Const CB_ERR = (-1)
Public Const CB_GETCOUNT = &H146
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158

Public Const CBEIF_TEXT = &H1
Public Const CBEIF_IMAGE = &H2
Public Const CBEIF_SELECTEDIMAGE = &H4

Public Const CBEM_SETITEMA = (WM_USER + 5)
Public Const CBEM_GETITEMA = (WM_USER + 4)
Public Const CBEM_DELETEITEM = &H144
Public Const CBEM_INSERTITEMA = (WM_USER + 1)

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

    strTempDir = left$(strBuffer, lngRet)

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
        Text = left$(Text, lngPos - 1)
    End If

    TrimNull = Text

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-17 19:37)  Decl: 184  Code: 75  Total: 259 Lines
':) CommentOnly: 11 (4.2%)  Commented: 1 (0.4%)  Filled: 187 (72.2%)  Empty: 72 (27.8%)  Max Logic Depth: 2
