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
        .hwnd = m_hWndParent
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

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-17 19:42)  Decl: 9  Code: 96  Total: 105 Lines
':) CommentOnly: 2 (1.9%)  Commented: 0 (0%)  Filled: 66 (62.9%)  Empty: 39 (37.1%)  Max Logic Depth: 2
