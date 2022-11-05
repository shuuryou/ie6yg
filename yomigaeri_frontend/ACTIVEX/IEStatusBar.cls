VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements SSubTUP.ISubclass

Private Const STATUS_PANES As Integer = 5
Private Const STATUS_PANE_NAVIGATION As Integer = 0
Private Const STATUS_PANE_PROGRESS As Integer = 1
Private Const STATUS_PANE_CONNECTION As Integer = 2
Private Const STATUS_PANE_SSL As Integer = 3
Private Const STATUS_PANE_ZONE As Integer = 4

Private Const ICON_ID_DISCONNECTED As Integer = 101
Private Const ICON_ID_CONNECTED As Integer = 102
Private Const ICON_ID_CONNECTING As Integer = 107
Private Const ICON_ID_WEBPAGE As Integer = 103
Private Const ICON_ID_SSL_BAD As Integer = 104
Private Const ICON_ID_SSL_OK As Integer = 105
Private Const ICON_ID_ZONE As Integer = 106

Private m_hWndIEFrame As Long
Private m_hWndStatusBar As Long
Private m_hWndProgressBar As Long

Private m_hIconDisconnected As Long
Private m_hIconConnected As Long
Private m_hIconConnecting As Long
Private m_hIconWebpage As Long
Private m_hIconSSLBad As Long
Private m_hIconSSLOK As Long
Private m_hIconZone As Long

Private m_EngineText() As Byte

Public Enum ConnectionIconState
    None = 0
    Disconnected = 1
    Connected = 2
    Connecting = 3
End Enum
#If False Then ':) Line inserted by Formatter
Private None, Disconnected, Connected, Connecting ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Public Enum SSLIconState
    None = 0
    OK = 1
    Bad = 2
End Enum
#If False Then ':) Line inserted by Formatter
Private None, OK, Bad ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Private m_CurrentStatusText As String
Private m_StateConnectionIcon As ConnectionIconState
Private m_StateSSLIcon As SSLIconState

Private m_IsStatusBarCurrentlySimple As Boolean

Public Property Let ConnectionIcon(newState As ConnectionIconState)

  Dim lngIcon As Long

    m_StateConnectionIcon = newState

    If m_StateConnectionIcon = Connected Then
        lngIcon = m_hIconConnected
      ElseIf m_StateConnectionIcon = Disconnected Then
        lngIcon = m_hIconDisconnected
      ElseIf m_StateConnectionIcon = Connecting Then
        lngIcon = m_hIconConnecting
      Else
        lngIcon = 0&
    End If

    SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, STATUS_PANE_CONNECTION, ByVal lngIcon

End Property

Public Property Get ConnectionIcon() As ConnectionIconState

    ConnectionIcon = m_StateConnectionIcon

End Property

Public Sub Construct(hWndIEFrame As Long)

  ' VB6 classes can't have a real constructor. *sigh*

  Dim lngIconSize As Long

    m_hWndIEFrame = hWndIEFrame

    FindStatusBar

    If m_hWndStatusBar = -1 Or m_hWndProgressBar = -1 Then
        Exit Sub
    End If

    m_EngineText = StrConv(LoadResString(106), vbFromUnicode) & Chr$(0) ' Chromium

    AttachMessage Me, m_hWndStatusBar, SB_SETTEXTA
    AttachMessage Me, m_hWndStatusBar, SB_SETTEXTW
    AttachMessage Me, m_hWndStatusBar, SB_SETPARTS
    AttachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT
    AttachMessage Me, m_hWndStatusBar, SB_SETICON
    AttachMessage Me, m_hWndStatusBar, SB_SIMPLE

    AttachMessage Me, m_hWndStatusBar, WM_SIZE
    AttachMessage Me, m_hWndStatusBar, WM_WINDOWPOSCHANGING
    AttachMessage Me, m_hWndStatusBar, WM_WINDOWPOSCHANGED

    AttachMessage Me, m_hWndStatusBar, SB_SETTEXTA + SB_IE6YG_SHIFT
    AttachMessage Me, m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT
    AttachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT + SB_IE6YG_SHIFT
    AttachMessage Me, m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT

    lngIconSize = GetSystemMetrics(SM_CXSMICON)

    m_hIconDisconnected = LoadImage(App.hInstance, MAKEINTRESOURCE(ICON_ID_DISCONNECTED), IMAGE_ICON, lngIconSize, lngIconSize, 0&)
    m_hIconConnected = LoadImage(App.hInstance, MAKEINTRESOURCE(ICON_ID_CONNECTED), IMAGE_ICON, lngIconSize, lngIconSize, 0&)
    m_hIconConnecting = LoadImage(App.hInstance, MAKEINTRESOURCE(ICON_ID_CONNECTING), IMAGE_ICON, lngIconSize, lngIconSize, 0&)
    m_hIconWebpage = LoadImage(App.hInstance, MAKEINTRESOURCE(ICON_ID_WEBPAGE), IMAGE_ICON, lngIconSize, lngIconSize, 0&)
    m_hIconSSLBad = LoadImage(App.hInstance, MAKEINTRESOURCE(ICON_ID_SSL_BAD), IMAGE_ICON, lngIconSize, lngIconSize, 0&)
    m_hIconSSLOK = LoadImage(App.hInstance, MAKEINTRESOURCE(ICON_ID_SSL_OK), IMAGE_ICON, lngIconSize, lngIconSize, 0&)
    m_hIconZone = LoadImage(App.hInstance, MAKEINTRESOURCE(ICON_ID_ZONE), IMAGE_ICON, lngIconSize, lngIconSize, 0&)

    FixStatusBar

End Sub

Public Sub Destroy()

    If m_hWndStatusBar = -1 Or m_hWndProgressBar = -1 Then
        Exit Sub
    End If

  Dim nParts As Long
  Dim n As Integer

    nParts = SendMessage(m_hWndStatusBar, SB_GETPARTS, 0, 0&)

    For n = 0 To nParts - 1 Step 1
        SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, n, 0&
    Next n

    DestroyIcon m_hIconDisconnected
    DestroyIcon m_hIconConnected
    DestroyIcon m_hIconWebpage
    DestroyIcon m_hIconSSLBad
    DestroyIcon m_hIconSSLOK
    DestroyIcon m_hIconZone

    DetachMessage Me, m_hWndStatusBar, SB_SETTEXTA
    DetachMessage Me, m_hWndStatusBar, SB_SETTEXTW
    DetachMessage Me, m_hWndStatusBar, SB_SETPARTS
    DetachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT
    DetachMessage Me, m_hWndStatusBar, SB_SETICON
    DetachMessage Me, m_hWndStatusBar, SB_SIMPLE

    DetachMessage Me, m_hWndStatusBar, WM_SIZE
    DetachMessage Me, m_hWndStatusBar, WM_WINDOWPOSCHANGING
    DetachMessage Me, m_hWndStatusBar, WM_WINDOWPOSCHANGED

    DetachMessage Me, m_hWndStatusBar, SB_SETTEXTA + SB_IE6YG_SHIFT
    DetachMessage Me, m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT
    DetachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT + SB_IE6YG_SHIFT
    DetachMessage Me, m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT

End Sub

Private Sub FindStatusBar()

  Dim hWndTemp As Long

    If m_hWndIEFrame = -1 Then
        modLogging.WriteLineToLog "IEStatusBar: m_hWndIEFrame is bad."
        m_hWndStatusBar = -1
        Exit Sub
    End If

    hWndTemp = FindWindowEx(m_hWndIEFrame, 0&, "msctls_statusbar32", vbNullString)

    If hWndTemp = 0 Then
        modLogging.WriteLineToLog "IEStatusBar: msctls_statusbar32 not found."
        Exit Sub
    End If

    m_hWndStatusBar = hWndTemp

    hWndTemp = FindWindowEx(hWndTemp, 0&, "msctls_progress32", vbNullString)

    If hWndTemp = 0 Then
        modLogging.WriteLineToLog "IEStatusBar: ReBarWindow32 not found."
        m_hWndStatusBar = -1
        Exit Sub
    End If

    m_hWndProgressBar = hWndTemp

    modLogging.WriteLineToLog "IEStatusBar: Found msctls_statusbar32: " & Hex$(m_hWndStatusBar) & " and msctls_progress32: " & Hex$(m_hWndProgressBar)

End Sub

Public Sub FixStatusBar()

    If m_hWndStatusBar = -1 Or m_hWndProgressBar = -1 Then
        Exit Sub
    End If

    ' Following code adapted from leaked WinXP code available at:
    ' http://<removed>/download/developer/xpsource/XPSP1/NT/shell/shdocvw/dochost.cpp
    ' void CDocObjectHost::_PlaceProgressBar(BOOL fForcedLayout)

    ' SB_IE6YG_SHIFT is used so WndProc can differentiate the
    ' messages from IE6YG and IE, because obviously IE will
    ' fight IE6YG to the grave to reclaim control over its
    ' status bar.

    SendMessage m_hWndStatusBar, SB_SIMPLE + SB_IE6YG_SHIFT, Abs(False), 0&

  Dim rc As RECT

    GetClientRect m_hWndStatusBar, rc

  Dim cxZone As Long
    cxZone = ZoneComputePaneSize()

  Dim cxProgressBar As Long
    cxProgressBar = 100

  Dim nSBWidth As Integer
    nSBWidth = rc.Right - rc.Left

  Dim arnRtEdge(STATUS_PANES) As Long
    arnRtEdge(0) = 1

  Dim nIconPaneWidth As Long
    nIconPaneWidth = GetSystemMetrics(SM_CXSMICON) + _
                     GetSystemMetrics(SM_CXEDGE) * 2

  Dim nWidthReqd As Long
    nWidthReqd = cxZone + cxProgressBar + (nIconPaneWidth * 1)

    ' Oh lord, VB6 not even has Max(). What a garbage language.

    arnRtEdge(STATUS_PANE_NAVIGATION) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    nWidthReqd = nWidthReqd - cxProgressBar
    arnRtEdge(STATUS_PANE_PROGRESS) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    nWidthReqd = nWidthReqd - nIconPaneWidth
    arnRtEdge(STATUS_PANE_CONNECTION) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    nWidthReqd = nWidthReqd - nIconPaneWidth
    arnRtEdge(STATUS_PANE_SSL) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    arnRtEdge(STATUS_PANE_ZONE) = -1

  Dim nParts As Long
    nParts = SendMessage(m_hWndStatusBar, SB_GETPARTS, 0, 0&)

    If nParts <> STATUS_PANES Then
  Dim n As Integer
        For n = 0 To nParts - 1 Step 1
            SendMessage m_hWndStatusBar, SB_SETTEXTA + SB_IE6YG_SHIFT, n Or SBT_NOTABPARSING, 0&
            SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, n, 0&
        Next n

        SendMessage m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT, 0, 0&
    End If

    SendMessage m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT, STATUS_PANES, arnRtEdge(0)

    If SendMessage(m_hWndStatusBar, SB_GETRECT, 1, rc) <> 0 Then
        InflateRect rc, -1 * GetSystemMetrics(SM_CXEDGE), -1 * GetSystemMetrics(SM_CYEDGE)
      Else
        rc.Left = rc.Top = rc.Right = rc.Bottom = 0
    End If

    SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, STATUS_PANE_NAVIGATION, ByVal m_hIconWebpage
    SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, STATUS_PANE_ZONE, ByVal m_hIconZone

    SendMessage m_hWndStatusBar, SB_SETMINHEIGHT + SB_IE6YG_SHIFT, _
                GetSystemMetrics(SM_CYSMICON) + GetSystemMetrics(SM_CYBORDER) * 2, 0&

    SendMessage m_hWndStatusBar, SB_SETTEXTA + SB_IE6YG_SHIFT, STATUS_PANE_ZONE, m_EngineText(0)

    SetWindowPos m_hWndProgressBar, 0&, _
                 rc.Left, rc.Top, _
                 rc.Right - rc.Left, rc.Bottom - rc.Top, _
                 SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW

    SendMessage m_hWndProgressBar, PBM_SETRANGE, 0&, ByVal MAKELPARAM(0, 100)

End Sub

Public Property Get hWndProgressBar() As Long

    hWndProgressBar = m_hWndProgressBar

End Property

Public Property Get hWndStatusBar() As Long

    hWndStatusBar = m_hWndStatusBar

End Property

Private Property Get ISubclass_MsgResponse() As SSubTUP.EMsgResponse

    ISubclass_MsgResponse = emrConsume

End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTUP.EMsgResponse)

  ' Unused

End Property ':(�No executable Code

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If iMsg = SB_SIMPLE Then
        m_IsStatusBarCurrentlySimple = (wParam > 0)
        ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
        Exit Function
    End If

    If iMsg = WM_SIZE Or iMsg = WM_WINDOWPOSCHANGING Or iMsg = WM_WINDOWPOSCHANGED Then
        FixStatusBar
    End If

    If (iMsg = SB_SETTEXTA Or iMsg = SB_SETTEXTW Or iMsg = SB_SETPARTS _
       Or iMsg = SB_SETMINHEIGHT Or iMsg = SB_SETICON) Then
       
        If m_IsStatusBarCurrentlySimple Then
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
        End If
       
        Exit Function
    End If

    If iMsg = SB_SETTEXTA + SB_IE6YG_SHIFT Or iMsg = SB_SETPARTS + SB_IE6YG_SHIFT Or _
       iMsg = SB_SETMINHEIGHT + SB_IE6YG_SHIFT Or iMsg = SB_SETICON + SB_IE6YG_SHIFT Then
        iMsg = iMsg - SB_IE6YG_SHIFT
    End If

    ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)

End Function

Public Property Get ProgressBarValue() As Integer

  Dim lngRet As Long

    lngRet = SendMessage(m_hWndProgressBar, PBM_GETPOS, 0&, 0&)

    If lngRet < 0 Then
        lngRet = 0
    End If

    If lngRet > 100 Then
        lngRet = 100
    End If

    ProgressBarValue = lngRet

End Property

Public Property Let ProgressBarValue(ByVal newValue As Integer)

    If newValue < 0 Then
        newValue = 0
    End If

    If newValue > 100 Then
        newValue = 100
    End If

    If m_hWndProgressBar = -1 Then
        Exit Property
    End If

    FixStatusBar ' TODO: Why is this needed for it to work?

    SendMessage m_hWndProgressBar, PBM_SETPOS, newValue, 0&

End Property

Public Property Get SSLIcon() As SSLIconState

    ConnectionIcon = m_StateConnectionIcon

End Property

Public Property Let SSLIcon(newState As SSLIconState)

  Dim lngIcon As Long

    m_StateSSLIcon = newState

    If m_StateSSLIcon = OK Then
        lngIcon = m_hIconSSLOK
      ElseIf m_StateSSLIcon = Bad Then
        lngIcon = m_hIconSSLBad
      Else
        lngIcon = 0&
    End If

    SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, STATUS_PANE_SSL, ByVal lngIcon

End Property

Public Property Get Text() As String

  Dim lngLength As Long
  Dim strText As String

    If m_hWndProgressBar = -1 Then
        Text = ""
        Exit Property
    End If

    lngLength = SendMessage(m_hWndProgressBar, SB_GETTEXTLENGTHA, 0&, 0&)

    If lngLength = 0 Then
        Text = ""
        Exit Property
    End If

    strText = String$(lngLength + 1, vbNullChar)

    SendMessage m_hWndStatusBar, SB_GETTEXTA, ByVal (lngLength + 1), ByVal StrPtr(strText)

    Text = strText

End Property

Public Property Let Text(newText As String)

  Dim bytText() As Byte

    If m_hWndProgressBar = -1 Then
        Exit Property
    End If

    bytText = StrConv(newText, vbFromUnicode) & Chr$(0)

    SendMessage m_hWndStatusBar, SB_SETTEXTA + SB_IE6YG_SHIFT, 0&, bytText(0)

End Property

Private Function ZoneComputePaneSize() As Long

  ' TODO: Use something more sensible than 80

    ZoneComputePaneSize = 80 + _
                          GetSystemMetrics(SM_CXSMICON) + _
                          GetSystemMetrics(SM_CXVSCROLL) + _
                          GetSystemMetrics(SM_CXEDGE) * 4

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-03 00:39)  Decl: 55  Code: 399  Total: 454 Lines
':) CommentOnly: 15 (3,3%)  Commented: 8 (1,8%)  Filled: 308 (67,8%)  Empty: 146 (32,2%)  Max Logic Depth: 3