VERSION 5.00
Object = "{8C11EFA1-92C3-11D1-BC1E-00C04FA31489}#1.0#0"; "mstscax.dll"
Begin VB.UserControl ctlFrontend 
   BackColor       =   &H80000005&
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   ScaleHeight     =   2955
   ScaleWidth      =   3975
   Begin MSTSCLibCtl.MsRdpClient2 rdpClient 
      Height          =   2655
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Server          =   ""
      FullScreen      =   0   'False
      StartConnected  =   0
   End
End
Attribute VB_Name = "ctlFrontend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const VIRTUAL_CHANNEL_NAME As String = "BEFECOM"

Private m_DoInitialConnect As Boolean

Private m_IEAddressBar As IEAddressBar
Private WithEvents m_IEBrowser As IEBrowser
Attribute m_IEBrowser.VB_VarHelpID = -1
Private WithEvents m_IEFrame As IEFrame
Attribute m_IEFrame.VB_VarHelpID = -1
Private m_IEStatusBar As IEStatusBar
Private WithEvents m_IEToolbar As IEToolbar
Attribute m_IEToolbar.VB_VarHelpID = -1

Private Sub m_IEBrowser_NavigationIntercepted(destinationURL As String)

    If rdpClient.Connected <> True Then
        Exit Sub
    End If

    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "NAVIGATE:" & destinationURL

End Sub

Private Sub m_IEFrame_WindowResized()

  Dim bWasConnected As Boolean

    bWasConnected = rdpClient.Connected

    If bWasConnected Then
        rdpClient.Disconnect

        Do Until rdpClient.Connected = False
            DoEvents
        Loop
    End If

    rdpClient.Top = 0
    rdpClient.Left = 0
    rdpClient.Width = UserControl.Width
    rdpClient.Height = UserControl.Height
    rdpClient.DesktopWidth = rdpClient.Width / Screen.TwipsPerPixelX
    rdpClient.DesktopHeight = rdpClient.Height / Screen.TwipsPerPixelY

    If bWasConnected Or m_DoInitialConnect Then
        If bWasConnected Then
            m_IEStatusBar.SetText LoadResString(104)  ' Reconnecting to rendering engine...
        End If

        rdpClient.Connect
    End If

End Sub

Private Sub m_IEToolbar_ToolbarButtonPressed(command As ToolbarCommand)

    Select Case command
      Case ToolbarCommand.CommandBack
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNBACK"
      Case ToolbarCommand.CommandForward
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNFORW"
      Case ToolbarCommand.CommandStop
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNSTOP"
      Case ToolbarCommand.CommandRefresh
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNREFR"
      Case Else
        modLogging.WriteLineToLog "ToolbarButtonPressed: Unknown command ID."
    End Select

End Sub

Public Sub PerformRemoteRefresh()

    If rdpClient.Connected Then
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "REFRESH"
    End If

End Sub

Public Sub QueryStreamingAvailable()

    MsgBox LoadResString(105), vbInformation ' Not implemented.

End Sub

Private Sub rdpClient_OnChannelReceivedData(ByVal chanName As String, ByVal data As String)

  Dim intPos As Integer
  Dim strURL As String
  Dim strTitle As String

    If chanName <> VIRTUAL_CHANNEL_NAME Then
        Exit Sub
    End If

    modLogging.WriteLineToLog "OnChannelReceivedData for " & chanName & ": " & data

    ' STYLING: Send colors of the client running frontend to backend
    ' CURSORS: Send paths to custom mouse cursors that backend will load via drive sharing
    ' LANGLST: Gets IE Accept-Language language list from registry or "<EMPTY>" if unset
    ' ADDRESS: Set IE address bar text to content following after "ADDRESS"
    ' ADDHIST: Add to IE history following after "ADDHIST" in the format URL\tTitle
    ' VISIBLE: Makes the RDP client visible
    ' INVISIB: Makes the RDP client invisible
    ' BBACKON: Toolbar BACK button ON
    ' BBACKOF: Toolbar BACK button OFF
    ' BFORWON: Toolbar FORWARD button ON
    ' BFORWOF: Toolbar FORWARD button OFF
    ' BSTOPON: Toolbar STOP button ON
    ' BSTOPOF: Toolbar STOP button OFF
    ' BREFRON: Toolbar REFRESH button ON
    ' BREFROF: Toolbar REFRESH button OFF
    ' BHOMEON: Toolbar HOME button ON
    ' BHOMEOF: Toolbar HOME button OFF
    ' BMEDION: Toolbar MEDIA button ON
    ' BMEDIOF: Toolbar MEDIA button OFF
    ' PGTITLE: Set the title of the page to the content following after "PGTITLE"
    ' SSLICON: Make the SSL icon visible with OK state
    ' SSLICBD: Make the SSL icon visible with error state
    ' SSLICOF: Make the SSL icon invisible

    Select Case Left$(UCase$(data), 7)
      Case "STYLING"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.MakeStyling(UserControl.hDC)
      Case "CURSORS"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.GetCursors()
      Case "LANGLST"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.GetAcceptLanguage()
      Case "ADDRESS"
        If Len(data) < 8 Then
            modLogging.WriteLineToLog "Cannot set address because data is too short."
            Exit Sub
        End If

        m_IEAddressBar.SetText Mid$(data, 8)
      Case "ADDHIST"
        If Len(data) < 8 Then
            modLogging.WriteLineToLog "Cannot add history because data is too short."
            Exit Sub
        End If

        intPos = InStr(1, data, vbTab, vbBinaryCompare)

        If intPos = 0 Then
            modLogging.WriteLineToLog "Cannot set address because of bad data."
            Exit Sub
        End If

        strURL = Trim$(Mid$(data, 8, intPos - 8))
        strTitle = Trim$(Mid$(data, intPos + 1))

        If strURL = "" Then
            modLogging.WriteLineToLog "Cannot set address because there is no URL."
        End If

        ' IUrlHistory gracefully deals with an empty title by making up a sensible one.

        m_IEBrowser.PushIntoHistory strURL, strTitle
      Case "VISIBLE"
        rdpClient.Visible = True

      Case "INVISIB"
        rdpClient.Visible = False
      Case "BBACKON"

        m_IEToolbar.SetToolbarCommandState CommandBack, True
      Case "BBACKOF"

        m_IEToolbar.SetToolbarCommandState CommandBack, False
      Case "BFORWON"

        m_IEToolbar.SetToolbarCommandState CommandForward, True
      Case "BFORWOF"

        m_IEToolbar.SetToolbarCommandState CommandForward, False
      Case "BSTOPON"

        m_IEToolbar.SetToolbarCommandState CommandStop, True
      Case "BSTOPOF"

        m_IEToolbar.SetToolbarCommandState CommandStop, False
      Case "BREFRON"

        m_IEToolbar.SetToolbarCommandState CommandRefresh, True
      Case "BREFROF"

        m_IEToolbar.SetToolbarCommandState CommandRefresh, False

      Case "BHOMEON"

        m_IEToolbar.SetToolbarCommandState CommandHome, True
      Case "BHOMEOF"

        m_IEToolbar.SetToolbarCommandState CommandHome, False

      Case "BMEDION"

        m_IEToolbar.SetToolbarCommandState CommandMedia, True
      Case "BMEDIOF"

        m_IEToolbar.SetToolbarCommandState CommandMedia, False
      Case "PGTITLE"
        If Len(data) < 8 Then
            modLogging.WriteLineToLog "Cannot set address because data is too short."
            Exit Sub
        End If

        m_IEBrowser.SetTitle Mid$(data, 8)
      Case "SSLICON"

        m_IEStatusBar.SSLIcon = SSLIconState.OK
      Case "SSLICBD"

        m_IEStatusBar.SSLIcon = SSLIconState.Bad
      Case "SSLICOF"

        m_IEStatusBar.SSLIcon = SSLIconState.None
      Case Else
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "UNSUPPORTED"

        modLogging.WriteLineToLog "Sent UNSUPPORTED response."
    End Select

End Sub

Private Sub rdpClient_OnConnected()

    m_IEStatusBar.ConnectionIcon = ConnectionIconState.Connected

End Sub

Private Sub rdpClient_OnConnecting()

    m_IEStatusBar.ConnectionIcon = ConnectionIconState.Connecting

End Sub

Private Sub rdpClient_OnDisconnected(ByVal discReason As Long)

    m_IEStatusBar.SetText LoadResString(102)  ' Disconnected from rendering engine.
    m_IEStatusBar.ConnectionIcon = ConnectionIconState.Disconnected
    rdpClient.Visible = False

End Sub

Private Sub UserControl_Initialize()

    Set m_IEAddressBar = New IEAddressBar
    Set m_IEBrowser = New IEBrowser
    Set m_IEFrame = New IEFrame
    Set m_IEStatusBar = New IEStatusBar
    Set m_IEToolbar = New IEToolbar

    rdpClient.CreateVirtualChannels VIRTUAL_CHANNEL_NAME

    rdpClient.AdvancedSettings3.EnableAutoReconnect = True
    rdpClient.AdvancedSettings3.RedirectDrives = True
    rdpClient.AdvancedSettings3.RedirectPrinters = True
    rdpClient.AdvancedSettings3.EnableWindowsKey = False
    rdpClient.AdvancedSettings3.keepAliveInterval = 5000
    rdpClient.AdvancedSettings3.MaximizeShell = True
    rdpClient.AdvancedSettings3.PerformanceFlags = &H1F
    ' &H1F =
    ' TS_PERF_DISABLE_WALLPAPER |
    ' TS_PERF_DISABLE_FULLWINDOWDRAG |
    ' TS_PERF_DISABLE_MENUANIMATIONS |
    ' TS_PERF_DISABLE_THEMING |
    ' TS_PERF_ENABLE_ENHANCED
    rdpClient.ColorDepth = 24

    ' The RDP client stays hidden until backend is fully loaded and makes it visible.
    ' In VB6, this doesn't bother MsRdpClient6. It does in C# and it disconnects when
    ' Visible is False. *sigh*

    rdpClient.Visible = False

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    modLogging.ENABLE_DEBUG_LOG = CBool(PropBag.ReadProperty("DebugLog", False))

    modLogging.WriteLineToLog "--------------------------------------------------"

    rdpClient.Server = CStr(PropBag.ReadProperty("RDP_Server", vbNullString))
    rdpClient.UserName = CStr(PropBag.ReadProperty("RDP_Username", vbNullString))
    rdpClient.AdvancedSettings3.ClearTextPassword = CStr(PropBag.ReadProperty("RDP_Password", vbNullString))
    rdpClient.SecuredSettings2.StartProgram = CStr(PropBag.ReadProperty("RDP_Backend", vbNullString))

    If rdpClient.Server <> "" And rdpClient.UserName <> "" And CStr(PropBag.ReadProperty("RDP_Password", vbNullString)) <> "" And rdpClient.SecuredSettings2.StartProgram <> "" Then
        m_DoInitialConnect = True
      Else
        Err.Raise -1, "YOMIGAERI", LoadResString(103) ' The parameters for the frontend are incorrect.
    End If

End Sub

Private Sub UserControl_Show()

  ' Will be fired when the control is actually shown on the website.
  ' UserControl_Initialize still has the control floating in space.

    m_IEFrame.Constructor UserControl.hWnd
    m_IEAddressBar.Constructor m_IEFrame.hWndIEFrame
    m_IEBrowser.Constructor m_IEFrame.hWndInternetExplorerServer
    m_IEStatusBar.Constructor m_IEFrame.hWndIEFrame
    m_IEToolbar.Constructor m_IEFrame.hWndIEFrame

    m_IEBrowser.SetTitle ""

    m_IEToolbar.SetToolbarCommandState CommandBack, False
    m_IEToolbar.SetToolbarCommandState CommandForward, False
    m_IEToolbar.SetToolbarCommandState CommandHome, False
    m_IEToolbar.SetToolbarCommandState CommandMedia, False
    m_IEToolbar.SetToolbarCommandState CommandRefresh, False
    m_IEToolbar.SetToolbarCommandState CommandStop, False

    m_IEStatusBar.ConnectionIcon = ConnectionIconState.None
    m_IEStatusBar.SSLIcon = SSLIconState.None

    If m_DoInitialConnect Then
        m_IEStatusBar.SetText LoadResString(101)  ' Connecting to rendering engine...
    End If

End Sub

Private Sub UserControl_Terminate()

    Set m_IEAddressBar = Nothing
    Set m_IEBrowser = Nothing
    Set m_IEFrame = Nothing
    Set m_IEStatusBar = Nothing
    Set m_IEToolbar = Nothing

    DoEvents

End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-01 16:03)  Decl: 11  Code: 336  Total: 347 Lines
':) CommentOnly: 38 (11%)  Commented: 6 (1,7%)  Filled: 253 (72,9%)  Empty: 94 (27,1%)  Max Logic Depth: 3
