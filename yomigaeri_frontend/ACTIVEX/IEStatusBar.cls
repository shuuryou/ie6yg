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

Private Declare Function InflateRect Lib "USER32.DLL" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowPos Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Shift messages this amount to filter in WndProc
Private Const SB_IE6YG_SHIFT As Long = 255

Private Const SB_SETTEXTA As Long = WM_USER + 1
Private Const SB_GETTEXTA = WM_USER + 2
Private Const SB_GETTEXTLENGTHA = WM_USER + 3
Private Const SB_SETPARTS As Long = WM_USER + 4
Private Const SB_GETPARTS As Long = WM_USER + 6
Private Const SB_SETMINHEIGHT As Long = WM_USER + 8
Private Const SB_SIMPLE As Long = WM_USER + 9
Private Const SB_GETRECT As Long = WM_USER + 10
Private Const SB_SETTEXTW As Long = WM_USER + 11
Private Const SB_SETICON As Long = WM_USER + 15
Private Const SB_SETTIPTEXTA As Long = WM_USER + 16

Private Const SBT_NOTABPARSING As Long = &H800

Private Const PBM_GETPOS As Long = WM_USER + 8
Private Const PBM_SETPOS As Long = WM_USER + 2
Private Const PBM_SETRANGE As Long = WM_USER + 1

Private m_hWndStatusBar As Long
Private m_hWndProgressBar As Long

Private m_hIconDisconnected As Long
Private m_hIconConnected As Long
Private m_hIconConnecting As Long
Private m_hIconWebpage As Long
Private m_hIconSSLNone As Long
Private m_hIconSSLBad As Long
Private m_hIconSSLOK As Long
Private m_hIconZone As Long

Private m_EngineText() As Byte

Private m_EngineDisconnectedText() As Byte
Private m_EngineConnectedText() As Byte
Private m_EngineConnectingText() As Byte

Private m_SSLIconNoneText() As Byte
Private m_SSLIconOKText() As Byte
Private m_SSLIconBadText() As Byte

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
        SendMessage m_hWndStatusBar, SB_SETTIPTEXTA + SB_IE6YG_SHIFT, StatusBarPanes.Connection, m_EngineConnectedText(0)
      ElseIf m_StateConnectionIcon = Disconnected Then
        SendMessage m_hWndStatusBar, SB_SETTIPTEXTA + SB_IE6YG_SHIFT, StatusBarPanes.Connection, m_EngineDisconnectedText(0)
        lngIcon = m_hIconDisconnected
      ElseIf m_StateConnectionIcon = Connecting Then
        SendMessage m_hWndStatusBar, SB_SETTIPTEXTA + SB_IE6YG_SHIFT, StatusBarPanes.Connection, m_EngineConnectingText(0)
        lngIcon = m_hIconConnecting
      Else
        lngIcon = 0&
    End If

    SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, StatusBarPanes.Connection, ByVal lngIcon

End Property

Public Property Get ConnectionIcon() As ConnectionIconState

    ConnectionIcon = m_StateConnectionIcon

End Property

Public Sub Construct(hWndStatusBar As Long, hWndProgressBar As Long)

  ' VB6 classes can't have a real constructor. *sigh*

  Dim lngIconSize As Long

    m_hWndStatusBar = hWndStatusBar
    m_hWndProgressBar = hWndProgressBar

    If m_hWndStatusBar = -1 Or m_hWndProgressBar = -1 Then
        Exit Sub
    End If

    m_EngineText = StrConv(LoadResString(106), vbFromUnicode) & vbNullChar ' Chromium

    m_EngineDisconnectedText = StrConv(LoadResString(102), vbFromUnicode) & vbNullChar ' Disconnected from rendering engine.
    m_EngineConnectedText = StrConv(LoadResString(112), vbFromUnicode) & vbNullChar ' Connected to the rendering engine.
    m_EngineConnectingText = StrConv(LoadResString(101), vbFromUnicode) & vbNullChar ' Connecting to rendering engine...

    m_SSLIconNoneText = StrConv(LoadResString(111), vbFromUnicode) & vbNullChar ' Insecure connection
    m_SSLIconOKText = StrConv(LoadResString(109), vbFromUnicode) & vbNullChar ' Insecure connection: certificate validation failed
    m_SSLIconBadText = StrConv(LoadResString(110), vbFromUnicode) & vbNullChar ' Secure connection

    AttachMessage Me, m_hWndStatusBar, SB_SETTEXTA
    AttachMessage Me, m_hWndStatusBar, SB_SETTEXTW
    AttachMessage Me, m_hWndStatusBar, SB_SETPARTS
    AttachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT
    AttachMessage Me, m_hWndStatusBar, SB_SETICON
    AttachMessage Me, m_hWndStatusBar, SB_SIMPLE

    AttachMessage Me, m_hWndStatusBar, WM_SIZE
    AttachMessage Me, m_hWndStatusBar, WM_WINDOWPOSCHANGING
    AttachMessage Me, m_hWndStatusBar, WM_WINDOWPOSCHANGED
    AttachMessage Me, m_hWndStatusBar, WM_LBUTTONUP
    AttachMessage Me, m_hWndStatusBar, WM_RBUTTONUP

    AttachMessage Me, m_hWndStatusBar, SB_SETTEXTA + SB_IE6YG_SHIFT
    AttachMessage Me, m_hWndStatusBar, SB_SETTIPTEXTA + SB_IE6YG_SHIFT
    AttachMessage Me, m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT
    AttachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT + SB_IE6YG_SHIFT
    AttachMessage Me, m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT

    lngIconSize = GetSystemMetrics(SM_CXSMICON)

    m_hIconDisconnected = LoadImage(App.hInstance, MAKEINTRESOURCE(101), IMAGE_ICON, lngIconSize, lngIconSize, 0&)
    m_hIconConnected = LoadImage(App.hInstance, MAKEINTRESOURCE(102), IMAGE_ICON, lngIconSize, lngIconSize, 0&)
    m_hIconConnecting = LoadImage(App.hInstance, MAKEINTRESOURCE(107), IMAGE_ICON, lngIconSize, lngIconSize, 0&)

    m_hIconWebpage = LoadImage(App.hInstance, MAKEINTRESOURCE(103), IMAGE_ICON, lngIconSize, lngIconSize, 0&)

    m_hIconSSLNone = LoadImage(App.hInstance, MAKEINTRESOURCE(108), IMAGE_ICON, lngIconSize, lngIconSize, 0&)
    m_hIconSSLBad = LoadImage(App.hInstance, MAKEINTRESOURCE(104), IMAGE_ICON, lngIconSize, lngIconSize, 0&)
    m_hIconSSLOK = LoadImage(App.hInstance, MAKEINTRESOURCE(105), IMAGE_ICON, lngIconSize, lngIconSize, 0&)

    m_hIconZone = LoadImage(App.hInstance, MAKEINTRESOURCE(106), IMAGE_ICON, lngIconSize, lngIconSize, 0&)

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
    DetachMessage Me, m_hWndStatusBar, WM_LBUTTONUP
    DetachMessage Me, m_hWndStatusBar, WM_RBUTTONUP

    DetachMessage Me, m_hWndStatusBar, SB_SETTEXTA + SB_IE6YG_SHIFT
    DetachMessage Me, m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT
    DetachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT + SB_IE6YG_SHIFT
    DetachMessage Me, m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT

End Sub

Public Sub FixStatusBar()

    If m_hWndStatusBar = -1 Or m_hWndProgressBar = -1 Then
        Exit Sub
    End If

    ' Following code adapted from leaked WinXP code.

    ' File: shell/shdocvw/dochost.cpp
    ' Function: void CDocObjectHost::_PlaceProgressBar(BOOL fForcedLayout)

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

  Dim arnRtEdge(STATUSBAR_PANES) As Long
    arnRtEdge(0) = 1

  Dim nIconPaneWidth As Long
    nIconPaneWidth = GetSystemMetrics(SM_CXSMICON) + 2 + _
                     GetSystemMetrics(SM_CXEDGE) * 2

    ' 2 added to CXSMICON above to make icons look less cramped.
    ' Microsoft didn't do that though.

  Dim nWidthReqd As Long
    nWidthReqd = cxZone + cxProgressBar + (nIconPaneWidth * 1)

    ' Oh lord, VB6 not even has Max(). What a garbage language.

    arnRtEdge(StatusBarPanes.Navigation) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    nWidthReqd = nWidthReqd - cxProgressBar
    arnRtEdge(StatusBarPanes.Progress) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    nWidthReqd = nWidthReqd - nIconPaneWidth
    arnRtEdge(StatusBarPanes.Connection) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    nWidthReqd = nWidthReqd - nIconPaneWidth
    arnRtEdge(StatusBarPanes.SSL) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    arnRtEdge(StatusBarPanes.Zone) = -1

  Dim nParts As Long
    nParts = SendMessage(m_hWndStatusBar, SB_GETPARTS, 0, 0&)

    If nParts <> STATUSBAR_PANES Then
  Dim n As Integer
        For n = 0 To nParts - 1 Step 1
            SendMessage m_hWndStatusBar, SB_SETTEXTA + SB_IE6YG_SHIFT, n Or SBT_NOTABPARSING, 0&
            SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, n, 0&
        Next n

        SendMessage m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT, 0, 0&
    End If

    SendMessage m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT, STATUSBAR_PANES, arnRtEdge(0)

    If SendMessage(m_hWndStatusBar, SB_GETRECT, 1, rc) <> 0 Then
        InflateRect rc, -1 * GetSystemMetrics(SM_CXEDGE), -1 * GetSystemMetrics(SM_CYEDGE)
      Else
        rc.Left = rc.Top = rc.Right = rc.Bottom = 0
    End If

    SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, StatusBarPanes.Navigation, ByVal m_hIconWebpage
    SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, StatusBarPanes.Zone, ByVal m_hIconZone

    SendMessage m_hWndStatusBar, SB_SETMINHEIGHT + SB_IE6YG_SHIFT, _
                GetSystemMetrics(SM_CYSMICON) + GetSystemMetrics(SM_CYBORDER) * 2, 0&

    SendMessage m_hWndStatusBar, SB_SETTEXTA + SB_IE6YG_SHIFT, StatusBarPanes.Zone, m_EngineText(0)

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

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTUP.EMsgResponse)

  ' Unused

End Property

Private Property Get ISubclass_MsgResponse() As SSubTUP.EMsgResponse

    ISubclass_MsgResponse = emrConsume

End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If iMsg = WM_LBUTTONUP Or iMsg = WM_RBUTTONUP Then
        ' Stop IE from displaying its status bar popup menus.
        Exit Function
    End If

    If iMsg = SB_SIMPLE Then
        m_IsStatusBarCurrentlySimple = (wParam > 0)
        ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Exit Function
    End If

    If iMsg = WM_SIZE Or iMsg = WM_WINDOWPOSCHANGING Or iMsg = WM_WINDOWPOSCHANGED Then
        FixStatusBar
    End If

    If (iMsg = SB_SETTEXTA Or iMsg = SB_SETTEXTW Or iMsg = SB_SETPARTS _
        Or iMsg = SB_SETMINHEIGHT Or iMsg = SB_SETICON) Then

        If m_IsStatusBarCurrentlySimple Then
            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        End If

        Exit Function
    End If

    If iMsg = SB_SETTEXTA + SB_IE6YG_SHIFT Or iMsg = SB_SETPARTS + SB_IE6YG_SHIFT Or _
       iMsg = SB_SETMINHEIGHT + SB_IE6YG_SHIFT Or iMsg = SB_SETICON + SB_IE6YG_SHIFT Or _
       iMsg = SB_SETTIPTEXTA + SB_IE6YG_SHIFT Then
        iMsg = iMsg - SB_IE6YG_SHIFT
    End If

    ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)

End Function

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

    ' TODO: Why is this needed for it to work?
    If Not m_IsStatusBarCurrentlySimple Then
        FixStatusBar
    End If

    SendMessage m_hWndProgressBar, PBM_SETPOS, newValue, 0&

End Property

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

Public Property Get SSLIcon() As SSLIconState

    ConnectionIcon = m_StateConnectionIcon

End Property

Public Property Let SSLIcon(newState As SSLIconState)

  Dim lngIcon As Long

    m_StateSSLIcon = newState

    If m_StateSSLIcon = OK Then
        lngIcon = m_hIconSSLOK
        SendMessage m_hWndStatusBar, SB_SETTIPTEXTA + SB_IE6YG_SHIFT, StatusBarPanes.SSL, m_SSLIconOKText(0)
      ElseIf m_StateSSLIcon = Bad Then
        lngIcon = m_hIconSSLBad
        SendMessage m_hWndStatusBar, SB_SETTIPTEXTA + SB_IE6YG_SHIFT, StatusBarPanes.SSL, m_SSLIconBadText(0)
      Else
        lngIcon = m_hIconSSLNone
        SendMessage m_hWndStatusBar, SB_SETTIPTEXTA + SB_IE6YG_SHIFT, StatusBarPanes.SSL, m_SSLIconNoneText(0)
    End If

    SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, StatusBarPanes.SSL, ByVal lngIcon

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

    bytText = StrConv(newText, vbFromUnicode) & vbNullChar

    SendMessage m_hWndStatusBar, SB_SETTEXTA + SB_IE6YG_SHIFT, 0&, bytText(0)

End Property

Private Function ZoneComputePaneSize() As Long

  ' TODO: Implement text width calculation instead of using 80

    ZoneComputePaneSize = 80 + _
                          GetSystemMetrics(SM_CXSMICON) + _
                          GetSystemMetrics(SM_CXVSCROLL) + _
                          GetSystemMetrics(SM_CXEDGE) * 4

    ' The very first version of Microsoft Word for Windows was considered a
    ' "death march" project. It took forever. It kept slipping. The whole
    ' team was working ridiculous hours, the project was delayed again, and
    ' again, and again, and the stress was incredible.

    ' The story goes that one programmer, who had to write the code to
    ' calculate the height of a line of text, simply wrote "return 12;"
    ' and waited for the bug report to come in about how his function is not
    ' always correct.

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-18 16:14)  Decl: 74  Code: 420  Total: 494 Lines
':) CommentOnly: 26 (5.3%)  Commented: 13 (2.6%)  Filled: 343 (69.4%)  Empty: 151 (30.6%)  Max Logic Depth: 3
