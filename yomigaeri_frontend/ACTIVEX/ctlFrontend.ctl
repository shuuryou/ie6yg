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
   Begin VB.Menu mnuHistoryForward 
      Caption         =   "Forward"
      Begin VB.Menu mnuHistoryForwardItem 
         Caption         =   "hoge"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistoryForwardItem 
         Caption         =   "hoge"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistoryForwardItem 
         Caption         =   "hoge"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistoryForwardItem 
         Caption         =   "hoge"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistoryForwardItem 
         Caption         =   "hoge"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistoryForwardItem 
         Caption         =   "-"
         Index           =   99
      End
   End
   Begin VB.Menu mnuHistoryBack 
      Caption         =   "Back"
      Begin VB.Menu mnuHistoryBackItem 
         Caption         =   "hoge"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistoryBackItem 
         Caption         =   "hoge"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistoryBackItem 
         Caption         =   "hoge"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistoryBackItem 
         Caption         =   "hoge"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistoryBackItem 
         Caption         =   "hoge"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistoryBackItem 
         Caption         =   "-"
         Index           =   99
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

Private Sub m_IEToolbar_ToolbarMenuRequested(command As ToolbarCommand)

  ' Item 99 is only visible when there are no items in the history.
  ' See comment in SetHistoryMenu for the reasoning.

    If command = CommandBack Then
        If Not mnuHistoryBackItem(99).Visible Then
            PopupMenu mnuHistoryBack
          Else
            modLogging.WriteLineToLog "ToolbarMenuRequested: Dan't show it on Back button when empty."
        End If
        Exit Sub
    End If

    If command = CommandForward Then
        If Not mnuHistoryForwardItem(99).Visible Then
            PopupMenu mnuHistoryForward
          Else
            modLogging.WriteLineToLog "ToolbarMenuRequested: Dan't show it on Forward button when empty."
        End If
        Exit Sub
    End If

End Sub

Private Sub PositionRDPClient()

    rdpClient.Move _
                   IIf(m_HideRDP, -3000 * Screen.TwipsPerPixelX, 0), _
                   IIf(m_HideRDP, -3000 * Screen.TwipsPerPixelY, 0), _
                   UserControl.Width, _
                   UserControl.Height

    DoEvents

End Sub

Private Sub rdpClient_OnChannelReceivedData(ByVal chanName As String, ByVal data As String)

  Dim intPos As Integer

    If chanName <> VIRTUAL_CHANNEL_NAME Then
        Exit Sub
    End If

    modLogging.WriteLineToLog "OnChannelReceivedData: " & data

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
    ' M_CUTON: Enable menu item Edit>Cut
    ' M_CUTOF: Disable menu item Edit>Cut
    ' MCOPYON: Enable menu item Edit>Copy
    ' MCOPYOF: Disable menu item Edit>Copy
    ' MPASTON: Enable menu item Edit>Paste
    ' MPASTOF: Disable menu item Edit>Paste
    ' MINHIBK: Modify the history popup menu of the Back button (see implementation)
    ' MINHIFW: Modify the history popup menu of the Forward button (see implementation)

    Select Case Left$(UCase$(data), 7)
      Case "STYLING"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.MakeStyling(UserControl.hDC)
      Case "CURSORS"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.GetCursors()
      Case "LANGLST"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.GetAcceptLanguage()
      Case "ADDRESS"
        If Len(data) = 7 Then
            m_IEAddressBar.Text = vbNullString
            Exit Sub
        End If

        m_IEAddressBar.Text = Mid$(data, 8)
      Case "ADDHIST"
        If Len(data) < 8 Then
            modLogging.WriteLineToLog "OnChannelReceivedData: ADDHIST: Refuse because data is too short."
            Exit Sub
        End If

        intPos = InStr(1, data, Chr$(1), vbBinaryCompare)

        If intPos = 0 Then
            modLogging.WriteLineToLog "OnChannelReceivedData: ADDHIST: Refuse because of bad data."
            Exit Sub
        End If

  Dim values(0 To 1) As String
        values(0) = Trim$(Mid$(data, 8, intPos - 8))
        values(1) = Trim$(Mid$(data, intPos + 1))

        If values(0) = "" Then
            modLogging.WriteLineToLog "OnChannelReceivedData: ADDHIST: Refuse because there is no URL."
        End If

        ' IUrlHistory gracefully deals with an empty title by making up a sensible one.

        m_IEBrowser.AddToHistory values(1), values(0)
      Case "VISIBLE"
        m_HideRDP = False
        PositionRDPClient
      Case "INVISIB"
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
        m_IEFrame.MenuEditStopEnabled = True
      Case "BSTOPOF"
        m_IEToolbar.SetToolbarCommandState CommandStop, False
        m_IEFrame.MenuEditStopEnabled = True
      Case "BREFRON"
        m_IEToolbar.SetToolbarCommandState CommandRefresh, True
        m_IEFrame.MenuEditRefreshEnabled = True
      Case "BREFROF"
        m_IEToolbar.SetToolbarCommandState CommandRefresh, False
        m_IEFrame.MenuEditRefreshEnabled = False
      Case "BHOMEON"
        m_IEToolbar.SetToolbarCommandState CommandHome, True
      Case "BHOMEOF"
        m_IEToolbar.SetToolbarCommandState CommandHome, False
      Case "BMEDION"
        m_IEToolbar.SetToolbarCommandState CommandMedia, True
      Case "BMEDIOF"
        m_IEToolbar.SetToolbarCommandState CommandMedia, False
      Case "PGTITLE"
        If Len(data) = 7 Then
            m_IEBrowser.SetTitle vbNullString
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
            modLogging.WriteLineToLog "OnChannelReceivedData: PROGRES: Refuse because data is too short."
            Exit Sub
        End If

        On Error Resume Next
            values(0) = CInt(Mid$(data, 8))
            m_IEStatusBar.ProgressBarValue = values(0)
        On Error GoTo 0
      Case "STATUST"
        If Len(data) < 8 Then
            ' Show what IE would show when there's no status
            m_IEStatusBar.Text = LoadResString(108) ' Done
            Exit Sub
        End If
        m_IEStatusBar.Text = Mid$(data, 8)
      Case "M_CUTON":
        m_IEFrame.MenuEditCutEnabled = True
      Case "M_CUTOF":
        m_IEFrame.MenuEditCutEnabled = False
      Case "MCOPYON":
        m_IEFrame.MenuEditCopyEnabled = True
      Case "MCOPYOF":
        m_IEFrame.MenuEditCopyEnabled = False
      Case "MPASTON":
        m_IEFrame.MenuEditPasteEnabled = True
      Case "MPASTOF":
        m_IEFrame.MenuEditPasteEnabled = False
      Case "MINHIBK":
        If Len(data) < 8 Then
            modLogging.WriteLineToLog "OnChannelReceivedData: MINHIBK: Refuse because data is too short."
            Exit Sub
        End If

        SetHistoryMenu True, Mid$(data, 8)
      Case "MINHIFW":
        If Len(data) < 8 Then
            modLogging.WriteLineToLog "OnChannelReceivedData: MINHIFW: Refuse because data is too short."
            Exit Sub
        End If

        SetHistoryMenu False, Mid$(data, 8)

      Case Else
        modLogging.WriteLineToLog "OnChannelReceivedData: Unknown command ignored."
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

Private Sub SetHistoryMenu(back As Boolean, backend_data As String)

  Dim items() As String
  Dim i As Integer
  Dim gotone As Boolean

    items = Split(backend_data, Chr$(1), , vbBinaryCompare)

    If UBound(items) <> 4 Then
        ' {0, 1, 2, 3, 4}, so UBound is 4 in VB6, because of course it is.
        ' UBound is the maximum index of an array that is valid/accessible.
        ' It's technically not the item count, just always abused as such.

        modLogging.WriteLineToLog "Cannot set back menu because data is bad: " & UBound(items)
        Exit Sub
    End If

    ' items is now Title,Title,Title,Title,Title

    ' Item 99 is required because it is illegal in VB6 for all menu items to be
    ' set to Visible = False at the same time.

    If back Then
        mnuHistoryBackItem(99).Visible = True
      Else
        mnuHistoryForwardItem(99).Visible = True
    End If

    gotone = False

    For i = 0 To 4 Step 1
        If back Then
            mnuHistoryBackItem(i).Caption = items(i)
            mnuHistoryBackItem(i).Visible = Len(items(i)) > 0

            ' gotone = gotone or mnuHistoryBackItem(i).Visible
            ' will always evaluate to True. VB6 is painful. :(
            gotone = (gotone = True) Or (mnuHistoryBackItem(i).Visible = True)
          Else
            mnuHistoryForwardItem(i).Caption = items(i)
            mnuHistoryForwardItem(i).Visible = Len(items(i)) > 0

            gotone = (gotone = True) Or (mnuHistoryForwardItem(i).Visible = True)
        End If
    Next i

    If back Then
        mnuHistoryBackItem(99).Visible = Not gotone
      Else
        mnuHistoryForwardItem(99).Visible = Not gotone
    End If

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
    rdpClient.AdvancedSettings3.PerformanceFlags = &H9F
    ' &H9F =
    ' TS_PERF_DISABLE_WALLPAPER |
    ' TS_PERF_DISABLE_FULLWINDOWDRAG |
    ' TS_PERF_DISABLE_MENUANIMATIONS |
    ' TS_PERF_DISABLE_THEMING |
    ' TS_PERF_ENABLE_ENHANCED GRAPHICS |
    ' TS_PERF_ENABLE_FONT_SMOOTHING
    rdpClient.ColorDepth = 24

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

    m_IEFrame.Construct UserControl.hWnd
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

End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-11 04:01)  Decl: 12  Code: 488  Total: 500 Lines
':) CommentOnly: 59 (11.8%)  Commented: 10 (2%)  Filled: 396 (79.2%)  Empty: 104 (20.8%)  Max Logic Depth: 3
