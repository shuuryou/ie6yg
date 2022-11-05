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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Server          =   ""
      FullScreen      =   0   'False
      StartConnected  =   0
   End
   Begin VB.Menu HistoryMenuForward 
      Caption         =   "Forward"
      Begin VB.Menu HistoryMenuForwardItem 
         Caption         =   "Sorry"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu HistoryMenuForwardItem 
         Caption         =   "not"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu HistoryMenuForwardItem 
         Caption         =   "implemented"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu HistoryMenuForwardItem 
         Caption         =   "yet."
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu HistoryMenuForwardItem 
         Caption         =   "Lame!"
         Enabled         =   0   'False
         Index           =   5
      End
   End
   Begin VB.Menu HistoryMenuBack 
      Caption         =   "Back"
      Begin VB.Menu HistoryMenuBackItem 
         Caption         =   "Sorry"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu HistoryMenuBackItem 
         Caption         =   "not"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu HistoryMenuBackItem 
         Caption         =   "implemented"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu HistoryMenuBackItem 
         Caption         =   "yet."
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu HistoryMenuBackItem 
         Caption         =   "Lame!"
         Enabled         =   0   'False
         Index           =   5
      End
   End
End
Attribute VB_Name = "ctlFrontend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const VIRTUAL_CHANNEL_NAME As String = "BEFECOM"

Private m_IEAddressBar As IEAddressBar
Private WithEvents m_IEBrowser As IEBrowser
Attribute m_IEBrowser.VB_VarHelpID = -1
Private WithEvents m_IEFrame As IEFrame
Attribute m_IEFrame.VB_VarHelpID = -1
Private m_IEStatusBar As IEStatusBar
Private WithEvents m_IEToolbar As IEToolbar
Attribute m_IEToolbar.VB_VarHelpID = -1

Private m_DoInitialConnect As Boolean
Private m_HideRDP As Boolean

Private Sub m_IEBrowser_NavigationIntercepted(destinationURL As String)

    If rdpClient.Connected <> 1 Then
        Exit Sub
    End If

    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "NAVIGATE:" & destinationURL

    rdpClient.SetFocus

End Sub

Private Sub m_IEFrame_CommandReceived(command As IECommand)

    Select Case command
      Case IECommand.CommandEditCut
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "CLIPCUT"
      Case IECommand.CommandEditCopy
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "CLIPCPY"
      Case IECommand.CommandEditPaste
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "CLIPPST"
      Case IECommand.CommandEditRefresh
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNREFR"
      Case IECommand.CommandEditStop
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNSTOP"
      Case IECommand.CommandFavoritesAdd
        MsgBox "Not implemented yet.", vbInformation, "Lame!" ' XXX TODO
      Case IECommand.CommandFileProperties
        MsgBox "Not implemented yet.", vbInformation, "Lame!" ' XXX TODO
      Case IECommand.CommandViewSource
        MsgBox "Not implemented yet.", vbInformation, "Lame!" ' XXX TODO
      Case Else
        modLogging.WriteLineToLog "ToolbarButtonPressed: Unknown command ID."
    End Select

End Sub

Private Sub m_IEFrame_WindowResized()

  Dim bWasConnected As Boolean

    m_HideRDP = True
    PositionRDPClient

    bWasConnected = (rdpClient.Connected = 1)

    If bWasConnected Then
        rdpClient.Disconnect

        Do Until rdpClient.Connected = False
            DoEvents
        Loop
    End If

    If bWasConnected Then
        m_IEStatusBar.ConnectionIcon = ConnectionIconState.Connecting
        m_IEStatusBar.Text = LoadResString(104)  ' Reconnecting to rendering engine...
        Sleep 1000 ' XXX Fix this
        rdpClient.Connect
    End If

End Sub

Private Sub m_IEToolbar_ToolbarButtonPressed(command As ToolbarCommand)

    If rdpClient.Connected <> 1 Then
        Exit Sub
    End If

    Select Case command
      Case ToolbarCommand.CommandBack
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNBACK"
      Case ToolbarCommand.CommandForward
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNFORW"
      Case ToolbarCommand.CommandStop
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNSTOP"
      Case ToolbarCommand.CommandRefresh
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNREFR"
      Case ToolbarCommand.CommandHome
        MsgBox "Not implemented yet.", vbInformation, "Lame!"
      Case Else
        modLogging.WriteLineToLog "ToolbarButtonPressed: Unknown command ID."
    End Select

End Sub

Private Sub PositionRDPClient()

    rdpClient.Move _
                   IIf(m_HideRDP, -3000 * Screen.TwipsPerPixelX, 0), _
                   IIf(m_HideRDP, -3000 * Screen.TwipsPerPixelY, 0), _
                   UserControl.Width, _
                   UserControl.Height

    DoEvents

End Sub

Private Sub m_IEToolbar_ToolbarMenuRequested(command As ToolbarCommand)
    If command = CommandBack Then
        PopupMenu HistoryMenuBack
        Exit Sub
    End If
    
    If command = CommandForward Then
        PopupMenu HistoryMenuForward
        Exit Sub
    End If
End Sub

Private Sub rdpClient_OnChannelReceivedData(ByVal chanName As String, ByVal data As String)

  Dim intPos As Integer
  Dim Values(0 To 5) As Variant

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
    ' PROGRES: Set the value of the progress bar (value 0-100 follows after "PROGRES")
    ' STATUST: Set IE status bar text to content following after "STATUST"

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

        m_IEAddressBar.Text = Mid$(data, 8)
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

        Values(0) = Trim$(Mid$(data, 8, intPos - 8))
        Values(1) = Trim$(Mid$(data, intPos + 1))

        If Values(0) = "" Then
            modLogging.WriteLineToLog "Cannot set address because there is no URL."
        End If

        ' IUrlHistory gracefully deals with an empty title by making up a sensible one.

        m_IEBrowser.PushIntoHistory Values(0), Values(1)
      Case "VISIBLE"
        m_HideRDP = False
        PositionRDPClient

      Case "INVISIBLE"
        m_HideRDP = True
        PositionRDPClient

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
      Case "PROGRES"
        If Len(data) < 8 Then
            modLogging.WriteLineToLog "Cannot set progress bar value because data is too short."
            Exit Sub
        End If

        On Error Resume Next
            Values(0) = CInt(Mid$(data, 8))
            m_IEStatusBar.ProgressBarValue = Values(0)
        On Error GoTo 0
      Case "STATUST"
        If Len(data) < 8 Then
            modLogging.WriteLineToLog "Cannot set status because data is too short."
            Exit Sub
        End If

        m_IEStatusBar.Text = Mid$(data, 8)

      Case Else
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "UNSUPPORTED"

        modLogging.WriteLineToLog "Sent UNSUPPORTED response."
    End Select

End Sub

Private Sub rdpClient_OnConnected()

    m_IEStatusBar.ConnectionIcon = ConnectionIconState.Connected
    m_IEStatusBar.Text = LoadResString(107) ' Rendering engine found...

End Sub

Private Sub rdpClient_OnDisconnected(ByVal discReason As Long)

    m_IEStatusBar.Text = LoadResString(102)  ' Disconnected from rendering engine.
    m_IEStatusBar.ConnectionIcon = ConnectionIconState.Disconnected

    m_HideRDP = True
    PositionRDPClient

End Sub

Private Sub UserControl_EnterFocus()

    rdpClient.SetFocus

End Sub

Private Sub UserControl_Initialize()

    m_HideRDP = True
    PositionRDPClient

    Set m_IEAddressBar = New IEAddressBar
    Set m_IEBrowser = New IEBrowser
    Set m_IEFrame = New IEFrame
    Set m_IEStatusBar = New IEStatusBar
    Set m_IEToolbar = New IEToolbar

    rdpClient.CreateVirtualChannels VIRTUAL_CHANNEL_NAME

    rdpClient.AdvancedSettings3.EnableAutoReconnect = False
    rdpClient.AdvancedSettings3.RedirectDrives = True
    rdpClient.AdvancedSettings3.RedirectPrinters = True
    rdpClient.AdvancedSettings3.EnableWindowsKey = False
    rdpClient.AdvancedSettings3.keepAliveInterval = 5000
    rdpClient.AdvancedSettings3.MaximizeShell = False
    rdpClient.AdvancedSettings3.PerformanceFlags = &H1F
    ' &H1F =
    ' TS_PERF_DISABLE_WALLPAPER |
    ' TS_PERF_DISABLE_FULLWINDOWDRAG |
    ' TS_PERF_DISABLE_MENUANIMATIONS |
    ' TS_PERF_DISABLE_THEMING |
    ' TS_PERF_ENABLE_ENHANCED
    rdpClient.ColorDepth = 24

    'rdpClient.Visible = False

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

Private Sub UserControl_Resize()

    If m_DoInitialConnect Then
        m_DoInitialConnect = False

        ' This way it triggers just once when Trident has put the ActiveX
        ' control in its final place (filling up the document view).

        PositionRDPClient
        m_IEStatusBar.Text = LoadResString(101)  ' Connecting to rendering engine....
        rdpClient.Connect
    End If

End Sub

Private Sub UserControl_Show()

  ' Will be fired when the control is actually shown on the website.
  ' UserControl_Initialize still has the control floating in space.

    PositionRDPClient

    m_IEFrame.Construct UserControl.hwnd
    m_IEAddressBar.Construct m_IEFrame.hWndIEFrame
    m_IEBrowser.Construct m_IEFrame.hWndInternetExplorerServer
    m_IEStatusBar.Construct m_IEFrame.hWndIEFrame
    m_IEToolbar.Construct m_IEFrame.hWndIEFrame

    m_IEBrowser.SetTitle ""

    m_IEToolbar.SetToolbarCommandState CommandBack, False
    m_IEToolbar.SetToolbarCommandState CommandForward, False
    m_IEToolbar.SetToolbarCommandState CommandHome, True
    m_IEToolbar.SetToolbarCommandState CommandMedia, False
    m_IEToolbar.SetToolbarCommandState CommandRefresh, True
    m_IEToolbar.SetToolbarCommandState CommandStop, True

    m_IEStatusBar.ConnectionIcon = ConnectionIconState.None
    m_IEStatusBar.SSLIcon = SSLIconState.None

End Sub

Private Sub UserControl_Terminate()

    m_IEFrame.Destroy
    m_IEStatusBar.Destroy
    m_IEToolbar.Destroy
    m_IEBrowser.Destroy

    Set m_IEFrame = Nothing
    Set m_IEStatusBar = Nothing
    Set m_IEToolbar = Nothing
    Set m_IEBrowser = Nothing

    ' Otherwise IE will hang because IEFrame was subclassed
    ' and IE really doesn't like it even if its completely
    ' undone.

    'CoUninitialize
    'ExitProcess 0

End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-05 22:36)  Decl: 12  Code: 403  Total: 415 Lines
':) CommentOnly: 48 (11,6%)  Commented: 6 (1,4%)  Filled: 307 (74%)  Empty: 108 (26%)  Max Logic Depth: 3
