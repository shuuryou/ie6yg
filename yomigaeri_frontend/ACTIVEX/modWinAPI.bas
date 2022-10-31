Attribute VB_Name = "modWinAPI"
Option Explicit

Public Declare Function AccessibleObjectFromWindow Lib "oleacc.dll" (ByVal hWnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long
Public Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function InflateRect Lib "user32.dll" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal WParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type TBBUTTON
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    dwData As Long
    iString As Long
End Type

Public Const IIDSTR_IHTMLElement As String = "{3050f1ff-98b5-11cf-bb82-00aa00bdce0b}"
Public Const IIDSTR_IWebBrowser2 As String = "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"
Public Const IIDSTR_SWebBrowserApp As String = "{0002DF05-0000-0000-C000-000000000046}"

Public Const OBJID_CLIENT As Long = -4

Public Const GWL_WNDPROC As Long = -4

Public Const WM_SETTEXT As Long = &HC

Public Const WM_USER As Long = &H400
Public Const TB_BUTTONCOUNT As Long = (WM_USER + &H18)
Public Const TB_GETBUTTON As Long = (WM_USER + &H17)
Public Const TB_ENABLEBUTTON  As Long = (WM_USER + &H1)

' Shift messages this amount to filter in WndProc
Public Const SB_IE6YG_SHIFT As Long = 255

Public Const SB_SETTEXT As Long = WM_USER + 1
Public Const SB_SETPARTS As Long = WM_USER + 4
Public Const SB_GETPARTS As Long = WM_USER + 6
Public Const SB_SETMINHEIGHT As Long = WM_USER + 8
Public Const SB_SIMPLE As Long = WM_USER + 9
Public Const SB_GETRECT As Long = WM_USER + 10
Public Const SB_SETICON As Long = WM_USER + 15

Public Const SBT_NOTABPARSING As Long = &H800

Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_NOACTIVATE As Long = &H10
