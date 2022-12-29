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
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3480
      Top             =   2400
   End
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
   Begin VB.Menu mnuContextSelection 
      Caption         =   "Selection Context Menu"
      Begin VB.Menu mnuContextSelectionItem 
         Caption         =   "Cut"
         Index           =   0
      End
      Begin VB.Menu mnuContextSelectionItem 
         Caption         =   "&Copy"
         Index           =   1
      End
      Begin VB.Menu mnuContextSelectionItem 
         Caption         =   "&Paste"
         Index           =   2
      End
      Begin VB.Menu mnuContextSelectionItem 
         Caption         =   "Select &All"
         Index           =   3
      End
      Begin VB.Menu mnuContextSelectionItem 
         Caption         =   "Pr&int"
         Index           =   4
      End
   End
   Begin VB.Menu mnuContextDefault 
      Caption         =   "Default Context Menu"
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "&Back"
         Index           =   0
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "F&orward"
         Index           =   1
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "Select &All"
         Index           =   3
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "Create Shor&tcut"
         Index           =   5
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "Add to &Favorites..."
         Index           =   6
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "&View Source"
         Index           =   7
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "Pr&int"
         Index           =   9
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "&Refresh"
         Index           =   10
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuContextDefaultItem 
         Caption         =   "&Properties"
         Index           =   12
      End
   End
   Begin VB.Menu mnuContextImage 
      Caption         =   "Image Context Menu"
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "&Open Link"
         Index           =   0
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "Open Link in &New Window"
         Index           =   1
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "S&how Picture"
         Index           =   3
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "S&how Video"
         Index           =   4
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "S&how Audio"
         Index           =   5
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "&Save Picture As..."
         Index           =   6
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "&Save Video As..."
         Index           =   7
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "&Save Audio As..."
         Index           =   8
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "Set as Back&ground"
         Index           =   9
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "&Copy"
         Index           =   11
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "Copy Shor&tcut"
         Index           =   12
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "Add to &Favorites..."
         Index           =   14
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuContextImageItem 
         Caption         =   "P&roperties"
         Index           =   16
      End
   End
   Begin VB.Menu mnuContextAnchor 
      Caption         =   "Anchor Context Menu"
      Begin VB.Menu mnuContextAnchorItem 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnuContextAnchorItem 
         Caption         =   "Open in &New Window"
         Index           =   1
      End
      Begin VB.Menu mnuContextAnchorItem 
         Caption         =   "Save Target &As..."
         Index           =   2
      End
      Begin VB.Menu mnuContextAnchorItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextAnchorItem 
         Caption         =   "&Copy"
         Index           =   4
      End
      Begin VB.Menu mnuContextAnchorItem 
         Caption         =   "Copy Shor&tcut"
         Index           =   5
      End
      Begin VB.Menu mnuContextAnchorItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuContextAnchorItem 
         Caption         =   "Add to &Favorites..."
         Index           =   7
      End
      Begin VB.Menu mnuContextAnchorItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuContextAnchorItem 
         Caption         =   "P&roperties"
         Index           =   9
      End
   End
End
Attribute VB_Name = "ctlFrontend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements SSubTUP.ISubclass

Private Const VIRTUAL_CHANNEL_NAME As String = "BEFECOM"

Private m_IEAddressBar As IEAddressBar
Private WithEvents m_IEBrowser As IEBrowser
Attribute m_IEBrowser.VB_VarHelpID = -1
Private WithEvents m_IEFrame As IEFrame
Attribute m_IEFrame.VB_VarHelpID = -1
Private m_IEStatusBar As IEStatusBar
Private WithEvents m_IEToolbar As IEToolbar
Attribute m_IEToolbar.VB_VarHelpID = -1
Private m_IEToolTip As IEToolTip

Private WithEvents m_FindReplace As cFindReplace
Attribute m_FindReplace.VB_VarHelpID = -1
Private m_NewFind As Boolean

Private m_CertCurrentState As CertificateStates
Private m_CertTempFile As String

Private m_HideRDP As Boolean

Private m_DownloadServer As String
Private m_DownloadPort As String

Private m_ServerHostID As String
Private m_ClientHostID As String

Private Sub BackendUpdateWindowSize()
   
    If rdpClient.Connected <> 1 Then
        Exit Sub
    End If

    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "WINSIZE" & _
                                   (UserControl.Width / Screen.TwipsPerPixelX) & "," & _
                                   (UserControl.Height / Screen.TwipsPerPixelY)

End Sub

Private Sub HandleADDHIST(ByRef data As String)

  Dim strUrl As String
  Dim strTitle As String
  Dim intPos As Integer

    If Len(data) = 0 Then
        modLogging.WriteLineToLog "HandleADDHIST: Refuse because data is too short."
        Exit Sub
    End If

    intPos = InStr(1, data, Chr$(1), vbBinaryCompare)

    If intPos = 0 Then
        modLogging.WriteLineToLog "HandleADDHIST: Refuse because data is invalid."
        Exit Sub
    End If

    strTitle = Left$(data, intPos - 1)
    strUrl = Mid$(data, intPos + 1)

    If strUrl = "" Then
        modLogging.WriteLineToLog "HandleADDHIST: Refuse because there is no URL."
    End If

    ' IUrlHistory gracefully deals with an empty title by making up a sensible one.

    m_IEBrowser.AddToHistory strUrl, strTitle

End Sub

Private Sub HandleADDRESS(ByRef data As String)

    m_IEAddressBar.Address = data

End Sub

Private Sub HandleCERDATA(ByRef data As String)

  Dim bytCertData() As Byte
  Dim intFF As Integer

    If Len(data) = 0 Then
        modLogging.WriteLineToLog "HandleCERDATA: Clear old certificate."

        On Error GoTo EH
        intFF = FreeFile
        Open m_CertTempFile For Binary Access Write Lock Read Write As #intFF
        Close #intFF
        On Error GoTo 0

        modLogging.WriteLineToLog "HandleCERDATA: Cleared temp file: " & m_CertTempFile

        Exit Sub
    End If

    bytCertData = modBase64.Base64Decode(data)

    intFF = FreeFile

    On Error GoTo EH
    Open m_CertTempFile For Binary Access Write Lock Write As #intFF
    Put #intFF, 1, bytCertData
    Close #intFF
    On Error GoTo 0

    Erase bytCertData

    modLogging.WriteLineToLog "HandleCERDATA: Wrote out PEM file to: " & m_CertTempFile

Exit Sub

EH:
    modLogging.WriteLineToLog "HandleCERDATA: Internal error: " & Err.Number & " (" & Err.Description & ")"

    On Error Resume Next
        Kill m_CertTempFile
    On Error GoTo 0

End Sub

Private Sub HandleCERSHOW()

  Dim frmErrDlg As frmCertificateError

    Set frmErrDlg = New frmCertificateError

    frmErrDlg.CertificateState = m_CertCurrentState
    frmErrDlg.CertificateFile = m_CertTempFile
    frmErrDlg.ParentWindowHandle = m_IEFrame.hWndIEFrame
    frmErrDlg.Show vbModal

    If frmErrDlg.DialogResult = vbYes Then
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "CERTCALLBACK CONTINUE"
      Else
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "CERTCALLBACK CANCEL"
    End If

    Set frmErrDlg = Nothing

End Sub

Private Sub HandleCERSTAT(ByRef data As String)

    If Len(data) = 0 Then
        modLogging.WriteLineToLog "HandleCERSTAT: Failed because data is too short."
        Exit Sub
    End If

    On Error Resume Next
        m_CertCurrentState = CInt(data)
    On Error GoTo 0

    modLogging.WriteLineToLog "HandleCERSTAT: State becomes: " & m_CertCurrentState

End Sub

Private Sub HandleCURSORS()

    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.GetCursors()

End Sub

Private Sub HandleDWNLOAD(ByRef data As String)

  Dim strUrl As String

    strUrl = "http://" & m_DownloadServer & ":" & m_DownloadPort & "/dl/" & data

    modLogging.WriteLineToLog "HandleDWNLOAD: Download this URL: " & strUrl

    m_IEBrowser.DoFileDownload strUrl

End Sub

Private Sub HandleFILEUPL()

  Dim objDialog As New cCommonDialog

    With objDialog
        .CancelError = False
        .DialogTitle = LoadResString(500) ' Open
        .Filter = LoadResString(501)
        .hWnd = m_IEFrame.hWndIEFrame
        .flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST

        .ShowOpen

        If .FileName = "" Then
            Exit Sub
        End If

        modLogging.WriteLineToLog "HandleFILEUPL: Selected file is:" & .FileName

        rdpClient.SetFocus

        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "FILEUPLCALLBACK " & .FileName
    End With

    Set objDialog = Nothing

End Sub

Private Sub HandleGETINFO(ByRef data As String)

    Select Case UCase$(data)
      Case "STYLING"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.MakeStyling(UserControl.hDC)
      Case "LANGUAGES"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, modFrontendStyling.GetAcceptLanguage()
      Case "INITIALSIZE"
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, _
                                       (UserControl.Width / Screen.TwipsPerPixelX) & "," & _
                                       (UserControl.Height / Screen.TwipsPerPixelY)
      Case Else:
        modLogging.WriteLineToLog "HandleGETINFO: Refuse because of invalid argument."
    End Select

End Sub

Private Sub HandleJSDIALG(ByRef data As String)

  Dim intPos As Integer

  Dim intLenPrompt As Integer
  Dim intLenDefault As Integer

  Dim strType As String

  Dim strPrompt As String

  Dim strDefault As String
  Dim strResponse As String

  Dim Result As VbMsgBoxResult

    If Len(data) < 7 Then
        ' Shortest possible would be "ALERT "
        modLogging.WriteLineToLog "HandleJSDIALG: Failed because data is too short."
        Exit Sub
    End If

    intPos = InStr(1, data, " ", vbBinaryCompare)

    If intPos = 0 Then
        modLogging.WriteLineToLog "HandleJSDIALG: Refuse because data is invalid."
        Exit Sub
    End If

    strType = Left$(data, intPos - 1)
    strPrompt = Mid$(data, intPos + 1)

    Select Case UCase$(strType)
      Case "ALERT"
        Result = MsgBox(strPrompt, vbOKOnly Or vbExclamation, LoadResString(300))
      Case "CONFIRM"
        Result = MsgBox(strPrompt, vbOKCancel Or vbQuestion, LoadResString(300))
      Case "PROMPT"
        ' strPrompt has more data here. It looks like this:
        ' "00000005hello00000006world!"
        '  ^ prompt     ^ default

        On Error GoTo EH
        intLenPrompt = CInt(Left$(strPrompt, 8))
        intLenDefault = Mid$(strPrompt, 8 + intLenPrompt + 1, 8)
        On Error GoTo 0

        modLogging.WriteLineToLog "HandleJSDIALG: Parsed lengths: " & intLenPrompt & ", " & intLenDefault

        strDefault = Right$(strPrompt, intLenDefault)
        strPrompt = Mid$(strPrompt, 8 + 1, intLenPrompt)

        strResponse = InputBox(LoadResString(304) & vbCrLf & vbCrLf & strPrompt, LoadResString(303), strDefault)

        If Len(strResponse) <> 0 Then
            rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "JSCALLBACK OK " & strResponse
            Exit Sub
        End If

        ' TODO: BUG This is not how prompt() behaves. An empty response doesn't
        ' indicate Cancel was pressed; you can press OK with an empty response.
        ' But to fix it a custom InputBox has to be created. VB6 InputBox can't
        ' differentiate between OK and Cancel if the response is empty.
        Result = vbCancel
      Case "ONBEFOREUNLOAD"
        strPrompt = LoadResString(301) & vbCrLf & vbCrLf & strPrompt & vbCrLf & vbCrLf & LoadResString(302)

        Result = MsgBox(strPrompt, vbOKCancel Or vbExclamation, LoadResString(300))
      Case Else
        modLogging.WriteLineToLog "HandleJSDIALG: Refuse because type is invalid: " & strType
    End Select

    If Result = vbOK Then
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "JSCALLBACK OK"
      Else
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "JSCALLBACK CANCEL"
    End If

Exit Sub

EH:
    modLogging.WriteLineToLog "HandleJSDIALG: Protocol error. " & Err.Number & " " & Err.Description
    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "JSCALLBACK FAIL"

End Sub

Private Sub HandleMENUSET(ByRef data As String)

  Dim strCommand As String
  Dim blnState As Boolean
  Dim intPos As Integer

  ' MENUSET IDENTIFIER STATE

    intPos = InStr(1, data, " ", vbBinaryCompare) ' Space between IDENTIFIER and STATE

    If intPos = 0 Then
        modLogging.WriteLineToLog "HandleMENUSET: Refuse because separator not found."
        Exit Sub
    End If

    If intPos + 1 > Len(data) Then
        modLogging.WriteLineToLog "HandleMENUSET: Refuse because data is invalid."
        Exit Sub
    End If

    strCommand = UCase$(Left$(data, intPos - 1))
    blnState = (UCase$(Mid$(data, intPos + 1)) = "TRUE")

    Select Case strCommand
      Case "CUT"
        m_IEFrame.MenuEditCutEnabled = blnState
      Case "COPY"
        m_IEFrame.MenuEditCopyEnabled = blnState
      Case "PASTE"
        m_IEFrame.MenuEditPasteEnabled = blnState
      Case "REFRESH"
        m_IEFrame.MenuEditRefreshEnabled = blnState
      Case "STOP"
        m_IEFrame.MenuEditStopEnabled = blnState
      Case Else
        modLogging.WriteLineToLog "HandleMENUSET: Refuse because command """ & strCommand & """ is invalid."
    End Select

End Sub

Private Sub HandlePGTITLE(ByRef data As String)

    m_IEBrowser.PageTitle = data

End Sub

Private Sub HandlePROGRES(ByRef data As String)

  Dim intProgress As Integer

    If Len(data) = 0 Then
        modLogging.WriteLineToLog "HandlePROGRES: Refuse because data is too short."
        Exit Sub
    End If

    On Error Resume Next
        intProgress = CInt(data)
        m_IEStatusBar.ProgressBarValue = intProgress
    On Error GoTo 0

End Sub

Private Sub HandleSHOSTID(ByRef data As String)

    m_ServerHostID = data

End Sub

Private Sub HandleSSLICON(ByRef data As String)

    Select Case UCase$(data)
      Case "OK"
        m_IEStatusBar.SSLIcon = SSLIconState.OK
      Case "BAD"
        m_IEStatusBar.SSLIcon = SSLIconState.Bad
      Case "OFF"
        m_IEStatusBar.SSLIcon = SSLIconState.None
      Case Else
        modLogging.WriteLineToLog "HandleSSLICON: Refuse because data is invalid."
    End Select

End Sub

Private Sub HandleSTATUST(ByRef data As String)

    If Len(data) = 0 Then
        ' Show what IE would show when there's no status
        m_IEStatusBar.Text = LoadResString(108) ' Done
        Exit Sub
    End If

    m_IEStatusBar.Text = data

End Sub

Private Sub HandleTOOLBAR(ByRef data As String)

  Dim strCommand As String
  Dim blnState As Boolean
  Dim intPos As Integer

  ' TOOLBAR IDENTIFIER STATE

    intPos = InStr(1, data, " ", vbBinaryCompare) ' Space between IDENTIFIER and STATE

    If intPos = 0 Then
        modLogging.WriteLineToLog "HandleTOOLBAR: Refuse because separator not found."
        Exit Sub
    End If

    If intPos + 1 > Len(data) Then
        modLogging.WriteLineToLog "HandleTOOLBAR: Refuse because data is invalid."
        Exit Sub
    End If

    strCommand = UCase$(Left$(data, intPos - 1))
    blnState = (UCase$(Mid$(data, intPos + 1)) = "TRUE")

    Select Case strCommand
      Case "BACK"
        m_IEToolbar.SetToolbarCommandState CommandBack, blnState
      Case "FORWARD"
        m_IEToolbar.SetToolbarCommandState CommandForward, blnState
      Case "STOP"
        m_IEToolbar.SetToolbarCommandState CommandStop, blnState
      Case "REFRESH"
        m_IEToolbar.SetToolbarCommandState CommandRefresh, blnState
      Case "HOME"
        ' Ignored for now.
        'm_IEToolbar.SetToolbarCommandState CommandHome, blnState
      Case "MEDIA"
        ' Ignored for now.
        'm_IEToolbar.SetToolbarCommandState CommandMedia, blnState
      Case Else
        modLogging.WriteLineToLog "HandleTOOLBAR: Refuse because command """ & strCommand & """ is invalid."

    End Select

End Sub

Private Sub HandleTOOLTIP(ByRef data As String)

    If Len(data) = 0 Then
        m_IEToolTip.Visible = False
        Exit Sub
    End If

    m_IEToolTip.Text = data
    m_IEToolTip.Visible = True

End Sub

Private Sub HandleTRAVELLOG(back As Boolean, data As String)

  Dim items() As String
  Dim i As Integer
  Dim gotone As Boolean

    If Len(data) = 0 Then
        modLogging.WriteLineToLog "HandleTRAVELLOG: Refuse because data is too short."
        Exit Sub
    End If

    items = Split(data, Chr$(1), , vbBinaryCompare)

    If UBound(items) <> 4 Then
        ' {0, 1, 2, 3, 4}, so UBound is 4 in VB6, because of course it is.
        ' UBound is the maximum index of an array that is valid/accessible.
        ' It's technically not the item count, just always abused as such.

        modLogging.WriteLineToLog "HandleTRAVELLOG: Cannot set travel log because data is bad: " & UBound(items)
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

Private Sub HandleVISIBLE(ByRef data As String)

    m_HideRDP = (UCase$(data) <> "TRUE")
    PositionRDPClient

    m_IEAddressBar.Enabled = Not m_HideRDP

End Sub

Private Sub HandleWWWAUTH(ByRef data As String)

  Dim frmPassDlg As frmNetworkPassword
  Dim items() As String

    If Len(data) = 0 Then
        modLogging.WriteLineToLog "HandleWWWAUTH: Refuse because data is too short."
        Exit Sub
    End If

    items = Split(data, Chr$(1), , vbBinaryCompare)

    If UBound(items) <> 2 Then
        modLogging.WriteLineToLog "HandleWWWAUTH: Refuse because element count is wrong: " & UBound(items)
        Exit Sub
    End If

    Set frmPassDlg = New frmNetworkPassword

    frmPassDlg.SetHostID m_ServerHostID, m_ClientHostID
    frmPassDlg.Site = items(0)
    frmPassDlg.Realm = items(1)
    frmPassDlg.URL = items(2)
    frmPassDlg.GetCredentialsFromRegistry
    frmPassDlg.ParentWindowHandle = m_IEFrame.hWndIEFrame
    frmPassDlg.Show vbModal

    If frmPassDlg.DialogResult = vbOK Then
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "AUTHCALLBACK OK " & _
                                       frmPassDlg.Username & Chr$(1) & frmPassDlg.Password
      Else
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "AUTHCALLBACK CANCEL"
    End If

    Set frmPassDlg = Nothing

End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTUP.EMsgResponse)

  'Ignore

End Property

Private Property Get ISubclass_MsgResponse() As SSubTUP.EMsgResponse

    ISubclass_MsgResponse = emrConsume

End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    ISubclass_WindowProc = 1&

End Function

Private Sub m_FindReplace_DialogClosed()

    If rdpClient.Connected <> 1 Then
        Exit Sub
    End If

    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "FIND END"

End Sub

Private Sub m_FindReplace_FindNext(ByVal sToFind As String, ByVal eFlags As CommonDialogDirect6.EFindReplaceFlags)

  Dim strCommand As String

    If rdpClient.Connected <> 1 Then
        Exit Sub
    End If

    strCommand = "FIND " & IIf(m_NewFind, "1", "0") & " " & Format$(CLng(eFlags), "00000000") & " " & sToFind

    m_NewFind = False

    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, strCommand

End Sub

Private Sub m_IEBrowser_NavigationIntercepted(destinationURL As String)

    If rdpClient.Connected <> 1 Then
        Exit Sub
    End If

    m_FindReplace.CloseDialog

    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, _
                                   "NAVIGATE " & destinationURL

    rdpClient.SetFocus

End Sub

Private Sub m_IEFrame_CertificateErrorDialogRequested()

  Dim frmErrDlg As frmCertificateError

    If m_CertCurrentState = CertificateStates.None Then
        MsgBox LoadResString(220), vbOKOnly Or vbInformation, LoadResString(219)
        Exit Sub
    End If

    Set frmErrDlg = New frmCertificateError

    frmErrDlg.NoPromptMode = True
    frmErrDlg.CertificateState = m_CertCurrentState
    frmErrDlg.CertificateFile = m_CertTempFile
    frmErrDlg.ParentWindowHandle = m_IEFrame.hWndIEFrame
    frmErrDlg.Show vbModal

    Set frmErrDlg = Nothing

End Sub

Private Sub m_IEFrame_CommandReceived(command As IECommand)

    Select Case command
      Case IECommand.CommandEditCut
        MsgBox "Not implemented yet.", vbInformation, "Lame!" ' XXX TODO
      Case IECommand.CommandEditCopy
        MsgBox "Not implemented yet.", vbInformation, "Lame!" ' XXX TODO
      Case IECommand.CommandEditPaste
        MsgBox "Not implemented yet.", vbInformation, "Lame!" ' XXX TODO
      Case IECommand.CommandEditRefresh
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNREFR"
      Case IECommand.CommandEditStop
        rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "BTNSTOP"
      Case IECommand.CommandEditFind
        m_NewFind = True
        m_FindReplace.VBFindText m_IEFrame.hWndIEFrame, , FR_NOWHOLEWORD Or FR_DOWN
      Case IECommand.CommandFavoritesAdd
        m_IEBrowser.AddToFavorites m_IEFrame.hWndIEFrame, m_IEAddressBar.Address, m_IEBrowser.PageTitle
      Case IECommand.CommandFileProperties
        MsgBox "Not implemented yet.", vbInformation, "Lame!" ' XXX TODO
      Case IECommand.CommandViewSource
        MsgBox "Not implemented yet.", vbInformation, "Lame!" ' XXX TODO
      Case Else
        modLogging.WriteLineToLog "ToolbarButtonPressed: Unknown command ID."
    End Select

End Sub

Private Sub m_IEFrame_WindowResized()

    BackendUpdateWindowSize

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

Private Sub mnuHistoryBackItem_Click(index As Integer)

    If rdpClient.Connected <> 1 Then
        Exit Sub
    End If

    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "MNUBACK" & (index + 1)

End Sub

Private Sub mnuHistoryForwardItem_Click(index As Integer)

    If rdpClient.Connected <> 1 Then
        Exit Sub
    End If

    rdpClient.SendOnVirtualChannel VIRTUAL_CHANNEL_NAME, "MNUFORW" & (index + 1)

End Sub

Private Sub PositionRDPClient()

    rdpClient.Move _
                   IIf(m_HideRDP, -3000 * Screen.TwipsPerPixelX, 0), _
                   IIf(m_HideRDP, -3000 * Screen.TwipsPerPixelY, 0), _
                   rdpClient.Width, _
                   rdpClient.Height

End Sub

Private Sub rdpClient_OnChannelReceivedData(ByVal chanName As String, ByVal data As String)

  Dim intPos As Integer

    If chanName <> VIRTUAL_CHANNEL_NAME Then
        Exit Sub
    End If

    modLogging.WriteLineToLog "OnChannelReceivedData: " & data

    ' GETINFO: Get info about the client system
    ' ADDRESS: Set IE address bar text to the argument
    ' ADDHIST: Add to IE history in the format URL\tTitle
    ' VISIBLE: Makes the RDP client visible if argument is "TRUE", invisible otherwise
    ' TOOLBAR: Turn toolbar items on or off
    ' PGTITLE: Set the title of the page to the argument
    ' SSLICON: Set the SSL icon in the status bar
    ' PROGRES: Set the value of the progress bar (argument value 0-100)
    ' STATUST: Set IE status bar text to the argument
    ' MENUSET: Turn menu items on or off
    ' TRAVLBK: Modify the travel log of the Back button
    ' TRAVLFW: Modify the travel log of the Forward button
    ' TOOLTIP: Show a Win32 tooltip at mouse position with text in argument
    ' CERSTAT: Update SSL certificate status flags based on argument
    ' CERDATA: Retrieve SSL certificate from backend as PEM file
    ' CERSHOW: Show Security Alert prompt
    ' JSDIALG: Show a JavaScript dialog
    ' DWNLOAD: Trigger a download of the specified ID
    ' CURSORS: Send custom cursor settings from registry to backend
    ' SHOSTID: Sets the server Host ID for password saving
    ' WWWAUTH: Show password dialog for HTTP auth and send credentials to backend
    ' FILEUPL: Show an "Open" common dialog and relay path back to backend

  Dim strCommand As String
  Dim strData As String

    strCommand = Left$(UCase$(data), 7)
    strData = Mid$(data, 9) ' There's a space after the command

    Select Case strCommand
      Case "GETINFO"
        HandleGETINFO strData
      Case "ADDRESS"
        HandleADDRESS strData
      Case "ADDHIST"
        HandleADDHIST strData
      Case "VISIBLE"
        HandleVISIBLE strData
      Case "TOOLBAR"
        HandleTOOLBAR strData
      Case "PGTITLE"
        HandlePGTITLE strData
      Case "SSLICON"
        HandleSSLICON strData
      Case "PROGRES"
        HandlePROGRES strData
      Case "STATUST"
        HandleSTATUST strData
      Case "MENUSET"
        HandleMENUSET strData
      Case "TRAVLBK"
        HandleTRAVELLOG True, strData
      Case "TRAVLFW"
        HandleTRAVELLOG False, strData
      Case "TOOLTIP"
        HandleTOOLTIP strData
      Case "CERDATA"
        HandleCERDATA strData
      Case "CERSTAT"
        HandleCERSTAT strData
      Case "CERSHOW"
        HandleCERSHOW
      Case "JSDIALG"
        HandleJSDIALG strData
      Case "DWNLOAD"
        HandleDWNLOAD strData
      Case "CURSORS"
        HandleCURSORS
      Case "SHOSTID"
        HandleSHOSTID strData
      Case "WWWAUTH"
        HandleWWWAUTH strData
      Case "FILEUPL"
        HandleFILEUPL
      Case "CTXMENU"
        HandleCTXMENU strData
      Case Else
        modLogging.WriteLineToLog "OnChannelReceivedData: Unknown command ignored."
    End Select

End Sub

Private Sub HandleCTXMENU(ByRef data As String)

  Dim items() As String

    If Len(data) = 0 Then
        modLogging.WriteLineToLog "HandleCTXMENU: Refuse because data is too short."
        Exit Sub
    End If

    items = Split(data, " ", , vbBinaryCompare)

    ' For example: 263 159 DEFAULT 0VD 1VD 3VE 5VE 6VE 7VE 9VE 10VE 12VE
    ' Must always have the first three.

    If UBound(items) < 2 Then
        modLogging.WriteLineToLog "HandleCTXMENU: Refuse because element count is wrong: " & UBound(items)
        Exit Sub
    End If
    
    Dim intX As Integer
    Dim intY As Integer
    Dim strMenu As String
    
    On Error GoTo EH ' *sigh*
    intX = CInt(items(0))
    intY = CInt(items(1))
    On Error GoTo 0
    
    strMenu = UCase$(items(2))
    
    modLogging.WriteLineToLog "HandleCTXMENU: Request to open " & strMenu & " menu at coordinates X: " & intX & ", Y: " & intY
    
    ' I can't believe this works...
    Dim objMenu As Variant
    Dim objMenuItems As Variant
    
    Select Case strMenu
        Case "SELECTION"
            Set objMenu = mnuContextSelection
            Set objMenuItems = mnuContextSelectionItem
        Case "DEFAULT"
            Set objMenu = mnuContextDefault
            Set objMenuItems = mnuContextDefaultItem
        Case "IMAGE"
            Set objMenu = mnuContextImage
            Set objMenuItems = mnuContextImageItem
        Case "ANCHOR"
            Set objMenu = mnuContextAnchor
            Set objMenuItems = mnuContextAnchorItem
    End Select
    
    ' To get to the index we remove two flags and convert the final
    ' string to an integer. Not the most beautiful way, but it works.
     
    Dim intIndex As Integer
    Dim strItemData As String
    Dim blnVisible As Boolean
    Dim blnEnabled As Boolean
     
    Dim i As Integer
 
    For i = 3 To UBound(items) Step 1
        strItemData = items(i)
        
        blnEnabled = (Mid$(strItemData, Len(strItemData), 1) = "E")
        strItemData = Left$(strItemData, Len(strItemData) - 1)
        
        blnVisible = (Mid$(strItemData, Len(strItemData), 1) = "V")
        strItemData = Left$(strItemData, Len(strItemData) - 1)
        
        On Error GoTo EH
        intIndex = CInt(strItemData)
        On Error GoTo 0
        
        objMenuItems(intIndex).Enabled = blnEnabled
        objMenuItems(intIndex).Visible = blnVisible
        
        modLogging.WriteLineToLog "HandleCTXMENU: Interpreted """ & strItemData & """ as Index: " & intIndex & ", Enabled? " & blnEnabled & ", Visible? " & blnVisible
    Next i
    
    ' TODO: Convert X and Y to twips and actually use them? Not
    ' specifying X and Y makes VBRUN use the current mouse pos.
    PopupMenu objMenu
    
    Set objMenu = Nothing
    Set objMenuItems = Nothing
    
Exit Sub

EH:
    modLogging.WriteLineToLog "HandleCTXMENU: Parsing or conversion error: (" & Err.Number & ") " & Err.Description
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

Private Sub tmrResize_Timer()

    tmrResize.Enabled = False

    BackendUpdateWindowSize

End Sub

Private Sub UserControl_EnterFocus()

    rdpClient.SetFocus

End Sub

Private Sub UserControl_Initialize()

  Dim lngWidth As Long, lngHeight As Long
  Dim objRegistry As CRegistry

    lngWidth = GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX
    lngHeight = GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY

    rdpClient.Width = lngWidth
    rdpClient.Height = lngHeight

    m_HideRDP = True
    PositionRDPClient

    Set m_IEAddressBar = New IEAddressBar
    Set m_IEBrowser = New IEBrowser
    Set m_IEFrame = New IEFrame
    Set m_IEStatusBar = New IEStatusBar
    Set m_IEToolbar = New IEToolbar
    Set m_IEToolTip = New IEToolTip

    Set m_FindReplace = New cFindReplace

    rdpClient.CreateVirtualChannels VIRTUAL_CHANNEL_NAME

    rdpClient.AdvancedSettings3.EnableAutoReconnect = True
    rdpClient.AdvancedSettings.allowBackgroundInput = True
    rdpClient.AdvancedSettings3.RedirectDrives = True
    rdpClient.AdvancedSettings3.RedirectPrinters = True
    rdpClient.AdvancedSettings3.EnableWindowsKey = False
    rdpClient.AdvancedSettings3.AcceleratorPassthrough = True
    rdpClient.AdvancedSettings3.keepAliveInterval = 5000
    rdpClient.AdvancedSettings3.MaximizeShell = False
    rdpClient.AdvancedSettings3.DoubleClickDetect = True
    rdpClient.AdvancedSettings3.PerformanceFlags = &H190
    ' TS_PERF_ENABLE_ENHANCED GRAPHICS |
    ' TS_PERF_ENABLE_FONT_SMOOTHING |
    ' TS_PERF_ENABLE_DESKTOP_COMPOSITION
    rdpClient.ColorDepth = 24

    m_CertTempFile = TempName()
    On Error Resume Next
        Kill m_CertTempFile
    On Error GoTo 0
    m_CertTempFile = Replace$(UCase$(m_CertTempFile), ".TMP", ".CER")

    Set objRegistry = New CRegistry

    With objRegistry
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\IE6YG"

        If Not .KeyExists Then
            .CreateKey
        End If

        .Default = -1
        .ValueType = REG_SZ
        .ValueKey = "HostID"

        If VarType(.value) = vbString Then
            m_ClientHostID = .value
          Else
            m_ClientHostID = GetGUID_Plain
            .value = m_ClientHostID

            .ValueKey = "HostID_Warning"
            .value = "Do not change or delete HostID. Your saved passwords will be lost."
        End If
    End With

    Set objRegistry = Nothing

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  Dim blnBad As Boolean

    modLogging.ENABLE_DEBUG_LOG = CBool(PropBag.ReadProperty("DebugLog", False))

    modLogging.WriteLineToLog "--------------------------------------------------"

    rdpClient.Server = Trim$(CStr(PropBag.ReadProperty("RDP_Server", vbNullString)))
    rdpClient.AdvancedSettings3.RDPPort = CLng(PropBag.ReadProperty("RDP_Port", 3389))
    rdpClient.Username = Trim$(CStr(PropBag.ReadProperty("RDP_Username", vbNullString)))
    rdpClient.AdvancedSettings3.ClearTextPassword = CStr(PropBag.ReadProperty("RDP_Password", vbNullString))
    rdpClient.SecuredSettings2.StartProgram = Trim$(CStr(PropBag.ReadProperty("RDP_Shell", vbNullString)))

    m_DownloadServer = Trim$(CStr(PropBag.ReadProperty("Download_Server", vbNullString)))
    m_DownloadPort = CLng(PropBag.ReadProperty("Download_Port", -1))

    blnBad = rdpClient.Server = "" Or rdpClient.Username = "" Or _
             CStr(PropBag.ReadProperty("RDP_Password", vbNullString)) = "" Or _
             rdpClient.SecuredSettings2.StartProgram = "" Or _
             m_DownloadServer = "" Or m_DownloadPort = -1

    If blnBad Then
        Err.Raise -1, "YOMIGAERI", LoadResString(103) ' The parameters for the frontend are incorrect.
        Exit Sub
    End If

    rdpClient.Connect

End Sub

Private Sub UserControl_Resize()

    If rdpClient.Connected <> 1 Then
        Exit Sub
    End If

    tmrResize.Enabled = True

End Sub

Private Sub UserControl_Show()

  ' Will be fired when the control is actually shown on the website.
  ' UserControl_Initialize still has the control floating in space.

    SSubTUP.AttachMessage Me, UserControl.hWnd, WM_SETCURSOR

    m_IEFrame.Construct UserControl.hWnd
    m_IEAddressBar.Construct m_IEFrame.hWndIEFrame
    m_IEBrowser.Construct m_IEFrame.hWndInternetExplorerServer
    m_IEStatusBar.Construct m_IEFrame.hWndStatusBar, m_IEFrame.hWndProgressBar
    m_IEToolbar.Construct m_IEFrame.hWndIEFrame
    m_IEToolTip.Construct UserControl.hWnd

    m_IEAddressBar.Enabled = False
    m_IEBrowser.PageTitle = ""

    m_IEToolbar.SetToolbarCommandState CommandBack, False
    m_IEToolbar.SetToolbarCommandState CommandForward, False
    m_IEToolbar.SetToolbarCommandState CommandHome, True
    m_IEToolbar.SetToolbarCommandState CommandMedia, False
    m_IEToolbar.SetToolbarCommandState CommandRefresh, True
    m_IEToolbar.SetToolbarCommandState CommandStop, True

    m_IEStatusBar.ConnectionIcon = ConnectionIconState.None
    m_IEStatusBar.SSLIcon = SSLIconState.None

    m_IEStatusBar.Text = LoadResString(101)  ' Connecting to rendering engine....

End Sub

Private Sub UserControl_Terminate()

    SSubTUP.DetachMessage Me, UserControl.hWnd, WM_SETCURSOR

    m_IEFrame.Destroy
    m_IEStatusBar.Destroy
    m_IEToolbar.Destroy
    m_IEBrowser.Destroy
    m_IEToolTip.Destroy

    Set m_IEFrame = Nothing
    Set m_IEStatusBar = Nothing
    Set m_IEToolbar = Nothing
    Set m_IEBrowser = Nothing
    Set m_IEToolTip = Nothing

    Set m_FindReplace = Nothing

    On Error Resume Next
        Kill m_CertTempFile
    On Error GoTo 0

End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Dec-09 06:43)  Decl: 26  Code: 1023  Total: 1049 Lines
':) CommentOnly: 56 (5.3%)  Commented: 14 (1.3%)  Filled: 752 (71.7%)  Empty: 297 (28.3%)  Max Logic Depth: 3
