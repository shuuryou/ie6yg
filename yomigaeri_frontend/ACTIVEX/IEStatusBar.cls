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

Implements SSubTimer.ISubclass

Private Const STATUS_PANES As Integer = 4
Private Const STATUS_PANE_NAVIGATION As Integer = 0
Private Const STATUS_PANE_PROGRESS As Integer = 1
Private Const STATUS_PANE_SSL As Integer = 2
Private Const STATUS_PANE_ZONE As Integer = 3

Private m_hWndIEFrame As Long
Private m_hWndStatusBar As Long
Private m_hWndProgressBar As Long

Public Sub Constructor(hWndIEFrame As Long)
' VB6 classes can't have a real constructor. *sigh*
    m_hWndIEFrame = hWndIEFrame
    
    FindStatusBar
    
    If m_hWndStatusBar = -1 Or m_hWndProgressBar = -1 Then
        Exit Sub
    End If
    
    AttachMessage Me, m_hWndStatusBar, SB_SETTEXT
    AttachMessage Me, m_hWndStatusBar, SB_SETPARTS
    AttachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT
    AttachMessage Me, m_hWndStatusBar, SB_SETICON
    AttachMessage Me, m_hWndStatusBar, WM_USER + 11
    
    AttachMessage Me, m_hWndStatusBar, SB_SETTEXT + SB_IE6YG_SHIFT
    AttachMessage Me, m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT
    AttachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT + SB_IE6YG_SHIFT
    AttachMessage Me, m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT
    
    FixStatusBar
End Sub

Private Sub Class_Terminate()
    If m_hWndStatusBar = -1 Or m_hWndProgressBar = -1 Then
        Exit Sub
    End If
    
    DetachMessage Me, m_hWndStatusBar, SB_SETTEXT
    DetachMessage Me, m_hWndStatusBar, SB_SETPARTS
    DetachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT
    DetachMessage Me, m_hWndStatusBar, SB_SETICON
    DetachMessage Me, m_hWndStatusBar, WM_USER + 11
    
    DetachMessage Me, m_hWndStatusBar, SB_SETTEXT + SB_IE6YG_SHIFT
    DetachMessage Me, m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT
    DetachMessage Me, m_hWndStatusBar, SB_SETMINHEIGHT + SB_IE6YG_SHIFT
    DetachMessage Me, m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT
End Sub

Public Property Get hWndStatusBar() As Long
    hWndStatusBar = m_hWndStatusBar
End Property

Public Property Get hWndProgressBar() As Long
    hWndProgressBar = m_hWndProgressBar
End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)
    ' Unused
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse
    ISubclass_MsgResponse = emrConsume
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal WParam As Long, ByVal lParam As Long) As Long
    If iMsg = SB_SETTEXT Or iMsg = SB_SETPARTS Or iMsg = SB_SETMINHEIGHT _
        Or iMsg = SB_SETICON Or iMsg = WM_USER + 11 Then
        Exit Function
    End If

    If iMsg = SB_SETTEXT + SB_IE6YG_SHIFT Or iMsg = SB_SETPARTS + SB_IE6YG_SHIFT Or _
       iMsg = SB_SETMINHEIGHT + SB_IE6YG_SHIFT Or iMsg = SB_SETICON + SB_IE6YG_SHIFT Then
        iMsg = iMsg - SB_IE6YG_SHIFT
    End If
   
    CallOldWindowProc hWnd, iMsg, WParam, lParam
End Function

Private Sub FindStatusBar()

  Dim hWndTemp As Long

    If m_hWndIEFrame = -1 Then
    modLogging.WriteLineToLog "FindStatusBar: m_hWndIEFrame is bad."
        m_hWndStatusBar = -1
        Exit Sub '---> Bottom
    End If

    hWndTemp = FindWindowEx(m_hWndIEFrame, 0&, "msctls_statusbar32", vbNullString)

    If hWndTemp = 0 Then
        modLogging.WriteLineToLog "FindStatusBar: msctls_statusbar32 not found."
        Exit Sub '---> Bottom
    End If

    m_hWndStatusBar = hWndTemp

    hWndTemp = FindWindowEx(hWndTemp, 0&, "msctls_progress32", vbNullString)

    If hWndTemp = 0 Then
        modLogging.WriteLineToLog "FindStatusBar: ReBarWindow32 not found."
        m_hWndStatusBar = -1
        Exit Sub '---> Bottom
    End If

    m_hWndProgressBar = hWndTemp

    modLogging.WriteLineToLog "FindStatusBar: Found msctls_statusbar32: " & Hex$(m_hWndStatusBar) & " and msctls_progress32: " & Hex$(m_hWndProgressBar)

End Sub

Public Sub FixStatusBar()

    If m_hWndStatusBar = -1 Or m_hWndProgressBar = -1 Then
        Exit Sub '---> Bottom
    End If

    Rem Mark Off Silent

    ' Following code adapted from leaked WinXP code available at:
    ' http://<removed>/download/developer/xpsource/XPSP1/NT/shell/shdocvw/dochost.cpp
    ' void CDocObjectHost::_PlaceProgressBar(BOOL fForcedLayout)

    ' I found it by searching the Internet for window messages
    ' shown in Spy++ when IE6 navigates somewhere and updates
    ' the status bar.

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
                     GetSystemMetrics(SM_CXEDGE) * 1

  Dim nWidthReqd As Long
    nWidthReqd = cxZone + cxProgressBar + (nIconPaneWidth * 1)

    ' Oh lord, VB6 not even has Max(). What a garbage language.

    arnRtEdge(STATUS_PANE_NAVIGATION) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    nWidthReqd = nWidthReqd - cxProgressBar
    arnRtEdge(STATUS_PANE_PROGRESS) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    nWidthReqd = nWidthReqd - nIconPaneWidth
    arnRtEdge(STATUS_PANE_SSL) = IIf(1 > nSBWidth - nWidthReqd, 1, nSBWidth - nWidthReqd)

    arnRtEdge(STATUS_PANE_ZONE) = -1

  Dim nParts As Long
    nParts = SendMessage(m_hWndStatusBar, SB_GETPARTS, 0, 0&)

    If nParts <> STATUS_PANES Then
  Dim n As Integer
        For n = 0 To nParts - 1 Step 1
            SendMessage m_hWndStatusBar, SB_SETTEXT + SB_IE6YG_SHIFT, n Or SBT_NOTABPARSING, 0&
            SendMessage m_hWndStatusBar, SB_SETICON + SB_IE6YG_SHIFT, n, 0&
        Next n

        SendMessage m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT, 0, 0&
    End If

    SendMessage m_hWndStatusBar, SB_SETPARTS + SB_IE6YG_SHIFT, STATUS_PANES, arnRtEdge(0)

    If SendMessage(m_hWndStatusBar, SB_GETRECT, 1, rc) <> 0 Then
        InflateRect rc, -1 * GetSystemMetrics(SM_CXEDGE), -1 * GetSystemMetrics(SM_CYEDGE)
      Else 'NOT SENDMESSAGE(M_HWNDSTATUSBAR,...
        rc.Left = rc.Top = rc.Right = rc.Bottom = 0
    End If

    SetWindowPos m_hWndProgressBar, 0&, _
                 rc.Left, rc.Top, _
                 rc.Right - rc.Left, rc.Bottom - rc.Top, _
                 SWP_NOZORDER Or SWP_NOACTIVATE

    SendMessage m_hWndStatusBar, SB_SETMINHEIGHT + SB_IE6YG_SHIFT, _
                GetSystemMetrics(SM_CYSMICON) + GetSystemMetrics(SM_CYBORDER) * 2, 0&

  Dim text() As Byte
    text = StrConv(LoadResString(106), vbFromUnicode) & Chr$(0) ' IE6YG Engine

    SendMessage m_hWndStatusBar, SB_SETTEXT + SB_IE6YG_SHIFT, STATUS_PANE_ZONE, text(0)

    Rem Mark On

End Sub

Private Function ZoneComputePaneSize() As Long

  ' Essentially measures the width of a string and some
  ' cosmetic elements around it in shlwapi/security.cpp

  Dim s_cxMaxZoneText As Single

  ' No way to do this easily from *here*, so fuck it
  ' s_cxMaxZoneText = m_FrontendInstance.TextWidth(LoadResString(106)) ' IE6YG Engine

    s_cxMaxZoneText = 100

    ZoneComputePaneSize = s_cxMaxZoneText + _
                          GetSystemMetrics(SM_CXSMICON) + _
                          GetSystemMetrics(SM_CXVSCROLL) + _
                          GetSystemMetrics(SM_CXEDGE) * 4

End Function

Public Sub SetText(newText As String)
    Dim text() As Byte
    
    If m_hWndStatusBar = -1 Then
        modLogging.WriteLineToLog "SetStatusBarText: m_hWndStatusBar is bad."
        Exit Sub '---> Bottom
    End If

    text = StrConv(newText, vbFromUnicode) & Chr$(0)

    SendMessage m_hWndStatusBar, SB_SETTEXT + SB_IE6YG_SHIFT, STATUS_PANE_NAVIGATION, text(0)

    modLogging.WriteLineToLog "SetStatusBarText: Text becomes: " & newText
End Sub
