VERSION 5.00
Object = "{8C11EFA1-92C3-11D1-BC1E-00C04FA31489}#1.0#0"; "mstscax.dll"
Begin VB.UserControl ctlFrontend 
   BackColor       =   &H80000005&
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   ScaleHeight     =   2640
   ScaleWidth      =   3975
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin MSTSCLibCtl.MsRdpClient6NotSafeForScripting rdpClient 
      Height          =   2535
      Left            =   600
      OleObjectBlob   =   "ctlFrontend.ctx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   2895
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
Private WithEvents m_BrowserManager As IEBrowserManager
Attribute m_BrowserManager.VB_VarHelpID = -1

Private Sub Command1_Click()
    m_BrowserManager.SetStatusBarText "fuck"
End Sub

Private Sub m_BrowserManager_IEToolbarCommandClicked(commandId As Long)

    Select Case commandId
      Case m_BrowserManager.ToolbarIdCommandBack

        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNBACK"
      Case m_BrowserManager.ToolbarIdCommandForward

        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNFORW"
      Case m_BrowserManager.ToolbarIdCommandRefresh

        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNREFR"
      Case m_BrowserManager.ToolbarIdCommandStop

        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNSTOP"
      Case Else

        modLogging.WriteLineToLog "IEToolbarCommandClicked: Unknown command ID."
    End Select

End Sub

Private Sub m_BrowserManager_IEWantsToNavigate(newUrl As String)

    If rdpClient.Connected <> True Then
        Exit Sub '---> Bottom
    End If

    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "NAVIGATE:" & newUrl

End Sub

Private Sub rdpClient_OnChannelReceivedData(ByVal chanName As String, ByVal data As String)

  Dim intPos As Integer
  Dim strURL As String
  Dim strTitle As String

    If chanName <> VIRTUAL_CHANNEL_NAME Then
        Exit Sub '---> Bottom
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
    ' BMEDION: Toolbar MEDIA button ON
    ' BMEDIOF: Toolbar MEDIA button OFF

    Select Case Left$(UCase$(data), 7)
      Case "STYLING"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.MakeStyling(UserControl.hdc)
      Case "CURSORS"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.GetCursors()
      Case "LANGLST"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.GetAcceptLanguage()
      Case "ADDRESS"
        If Len(data) < 8 Then
            modLogging.WriteLineToLog "Cannot set address because data is too short."
            Exit Sub '---> Bottom
        End If

        m_BrowserManager.SetAddressBarText Mid$(data, 8)
      Case "ADDHIST"
        If Len(data) < 8 Then
            modLogging.WriteLineToLog "Cannot add history because data is too short."
            Exit Sub '---> Bottom
        End If

        intPos = InStr(1, data, vbTab, vbBinaryCompare)

        If intPos = 0 Then
            modLogging.WriteLineToLog "Cannot set address because of bad data."
            Exit Sub '---> Bottom
        End If

        strURL = Trim$(Mid$(data, 8, intPos - 8))
        strTitle = Trim$(Mid$(data, intPos + 1))

        If strURL = "" Then
            modLogging.WriteLineToLog "Cannot set address because there is no URL."
        End If

        ' IUrlHistory gracefully deals with an empty title by making up a sensible one.

        m_BrowserManager.PushIntoHistory strURL, strTitle
      Case "VISIBLE"
        rdpClient.Visible = True

      Case "INVISIB"
        rdpClient.Visible = False
      Case "BBACKON"

        m_BrowserManager.ToolbarButtonStateBack = True
      Case "BBACKOF"

        m_BrowserManager.ToolbarButtonStateBack = False
      Case "BFORWON"

        m_BrowserManager.ToolbarButtonStateForward = True
      Case "BFORWOF"

        m_BrowserManager.ToolbarButtonStateForward = False
      Case "BSTOPON"

        m_BrowserManager.ToolbarButtonStateStop = True
      Case "BSTOPOF"

        m_BrowserManager.ToolbarButtonStateStop = False
      Case "BREFRON"

        m_BrowserManager.ToolbarButtonStateRefresh = True
      Case "BREFROF"

        m_BrowserManager.ToolbarButtonStateRefresh = False

      Case Else
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "UNSUPPORTED"

        modLogging.WriteLineToLog "Sent UNSUPPORTED response."
    End Select

End Sub

Private Sub rdpClient_OnDisconnected(ByVal discReason As Long)

    m_BrowserManager.SetStatusBarText LoadResString(102)  ' Disconnected from rendering engine.
    rdpClient.Visible = False

End Sub

Private Sub UserControl_Initialize()

    Set m_BrowserManager = New IEBrowserManager
    Set modWndProc.BROWSER_MANAGER_INSTANCE = m_BrowserManager

    m_BrowserManager.hWndUserControl = UserControl.hWnd
    
    rdpClient.CreateVirtualChannels VIRTUAL_CHANNEL_NAME

    rdpClient.AdvancedSettings5.EnableAutoReconnect = True
    rdpClient.AdvancedSettings5.RedirectDrives = True
    rdpClient.AdvancedSettings5.RedirectPrinters = True
    rdpClient.AdvancedSettings5.EnableWindowsKey = False
    rdpClient.AdvancedSettings5.keepAliveInterval = 5000
    rdpClient.AdvancedSettings5.MaximizeShell = True
    rdpClient.AdvancedSettings5.PerformanceFlags = &H1F
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
    rdpClient.AdvancedSettings5.ClearTextPassword = CStr(PropBag.ReadProperty("RDP_Password", vbNullString))
    rdpClient.SecuredSettings2.StartProgram = CStr(PropBag.ReadProperty("RDP_Backend", vbNullString))

    If rdpClient.Server <> "" And rdpClient.UserName <> "" And CStr(PropBag.ReadProperty("RDP_Password", vbNullString)) <> "" And rdpClient.SecuredSettings2.StartProgram <> "" Then
        'm_DoInitialConnect = True 'XXX WRONG
      Else 'NOT RDPCLIENT.SERVER...
        Err.Raise -1, "YOMIGAERI", LoadResString(103) ' The parameters for the frontend are incorrect.
    End If

End Sub

Private Sub UserControl_Resize()

  Dim bWasConnected As Boolean

    m_BrowserManager.FixStatusBar

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
            m_BrowserManager.SetStatusBarText LoadResString(104)  ' Reconnecting to rendering engine...
        End If

        rdpClient.Connect
    End If

End Sub

Private Sub UserControl_Show()

  ' Will be fired when the control is actually shown on the website.
  ' UserControl_Initialize still has the control floating in space.

    m_BrowserManager.HookIWebBrowser2
    m_BrowserManager.HookButtonToolbarCommands
    m_BrowserManager.HookStatusBar
    m_BrowserManager.FixStatusBar

    If m_DoInitialConnect Then
        m_BrowserManager.SetStatusBarText LoadResString(101)  ' Connecting to rendering engine...
    End If

End Sub

Private Sub UserControl_Terminate()

    If Not m_BrowserManager Is Nothing Then
        m_BrowserManager.ReleaseIWebBrowser2
        m_BrowserManager.ReleaseButtonToolbarCommands
        m_BrowserManager.ReleaseStatusBar
    End If

    Set m_BrowserManager = Nothing

End Sub

Public Sub PerformRemoteRefresh()

    If rdpClient.Connected Then
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "REFRESH"
    End If

End Sub

Public Sub QueryStreamingAvailable()

    MsgBox LoadResString(105), vbInformation ' Not implemented.

End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Oct-29 22:35)  Decl: 8  Code: 268  Total: 276 Lines
':) CommentOnly: 32 (11.6%)  Commented: 6 (2.2%)  Filled: 195 (70.7%)  Empty: 81 (29.3%)  Max Logic Depth: 3
