VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateWindowEx Lib "USER32.DLL" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Private Declare Function GetCursorPos Lib "USER32.DLL" (lpPoint As POINT) As Long

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uid As Long
    rc As RECT
    hinst As Long
    lpszText As String
    lParam As Long
End Type

Private Const CW_USEDEFAULT As Long = &H80000000
Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"
Private Const HWND_TOPMOST As Long = -1

Private Const TTF_TRACK As Long = &H20
Private Const TTF_ABSOLUTE As Long = &H80

Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_NOPREFIX As Long = &H2

Private Const TTM_ADDTOOL As Long = (WM_USER + 4)
Private Const TTM_UPDATETIPTEXT As Long = (WM_USER + 12)
Private Const TTM_TRACKACTIVATE As Long = (WM_USER + 17)
Private Const TTM_TRACKPOSITION As Long = (WM_USER + 18)

Private m_HwndToolTip As Long
Private m_hWndParent As Long

Private m_CurrentText As String

Private m_CursorSize As POINT
Private m_ToolInfo As TOOLINFO

Private Sub Class_Initialize()

    m_hWndParent = -1
    m_HwndToolTip = -1

End Sub

Public Sub Construct(ByVal hWndParent As Long)

  Dim varRect As RECT

    m_CursorSize.X = GetSystemMetrics(SM_CXCURSOR)
    m_CursorSize.Y = GetSystemMetrics(SM_CYCURSOR)

    m_hWndParent = hWndParent

    m_HwndToolTip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASSA, vbNullString, WS_POPUP Or TTS_NOPREFIX Or TTS_ALWAYSTIP, _
                    CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, m_hWndParent, 0&, _
                    App.hInstance, ByVal 0&)

    If m_HwndToolTip = 0 Then
        modLogging.WriteLineToLog "IEToolTip: CreateWindowEx failed. HRESULT=" & Hex$(Err.LastDllError)
        m_HwndToolTip = -1
        Exit Sub
    End If

    GetClientRect m_hWndParent, varRect

    With m_ToolInfo
        .cbSize = Len(m_ToolInfo)
        .uFlags = TTF_ABSOLUTE Or TTF_TRACK
        .hWnd = m_hWndParent
        .uid = 1
        .lpszText = vbNullString
        .rc = varRect
    End With

    SendMessage m_HwndToolTip, TTM_ADDTOOL, 0&, m_ToolInfo

End Sub

Public Sub Destroy()

    If m_HwndToolTip = -1 Then
        Exit Sub
    End If

    DestroyWindow m_HwndToolTip

End Sub

Public Property Get Text() As String

    Text = m_CurrentText

End Property

Public Property Let Text(value As String)

  Dim text_cstr() As Byte

    If m_CurrentText = value Then
        Exit Property
    End If

    m_CurrentText = value

    Text = StrConv(m_CurrentText & vbNullChar, vbFromUnicode)

    m_ToolInfo.lpszText = text_cstr(0)

    SendMessage m_HwndToolTip, TTM_UPDATETIPTEXT, 0&, m_ToolInfo

End Property

Public Property Let Visible(value As Boolean)

  Dim varPoint As POINT
  Dim lngPosition As Long

    If m_HwndToolTip = -1 Then
        Exit Property
    End If

    GetCursorPos varPoint

    lngPosition = MAKELPARAM(varPoint.X, varPoint.Y + (m_CursorSize.Y / 2))

    SendMessage m_HwndToolTip, TTM_TRACKPOSITION, 0&, ByVal lngPosition
    SendMessage m_HwndToolTip, TTM_TRACKACTIVATE, IIf(value, 1&, 0&), m_ToolInfo

End Property

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-18 00:20)  Decl: 39  Code: 96  Total: 135 Lines
':) CommentOnly: 2 (1.5%)  Commented: 0 (0%)  Filled: 90 (66.7%)  Empty: 45 (33.3%)  Max Logic Depth: 2
