VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements SSubTUP.ISubclass

Private m_hWndUserControl As Long
Private m_hWndInternetExplorerServer As Long
Private m_hWndShellDocObjectView As Long
Private m_hWndIEFrame As Long
Private m_GotSizeMessage As Boolean
Private m_WasMaximized As Boolean
Private m_WasMinimized As Boolean

Private Const IE_MENU_COMMAND_FILE_PROPERTIES As Long = &H106
Private Const IE_MENU_COMMAND_EDIT_CUT As Long = &HA041
Private Const IE_MENU_COMMAND_EDIT_COPY As Long = &HA042
Private Const IE_MENU_COMMAND_EDIT_PASTE As Long = &HA043
Private Const IE_MENU_COMMAND_EDIT_STOP As Long = &HA21A
Private Const IE_MENU_COMMAND_EDIT_REFRESH As Long = &HA220
Private Const IE_MENU_COMMAND_EDIT_FIND As Long = &H443
Private Const IE_MENU_COMMAND_VIEW_SOURCE As Long = &HC5B
Private Const IE_MENU_COMMAND_FAVORITES_ADD As Long = &HA173

Public Enum IECommand
    CommandNone = 0
    CommandFileProperties = 1
    CommandEditCut = 2
    CommandEditCopy = 3
    CommandEditPaste = 4
    CommandEditStop = 5
    CommandEditRefresh = 6
    CommandEditFind = 7
    CommandViewSource = 8
    CommandFavoritesAdd = 9
End Enum
#If False Then ':) Line inserted by Formatter
Private CommandNone, CommandFileProperties, CommandEditCut, CommandEditCopy, CommandEditPaste, CommandEditStop, CommandEditRefresh, CommandEditFind, _
        CommandViewSource, CommandFavoritesAdd ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Private m_MenuEditCutEnabled As Boolean
Private m_MenuEditCopyEnabled As Boolean
Private m_MenuEditPasteEnabled As Boolean
Private m_MenuEditStopEnabled As Boolean
Private m_MenuEditRefreshEnabled As Boolean

Public Event WindowResized()
Public Event CommandReceived(command As IECommand)

Public Sub Construct(hWndUserControl As Long)

  ' VB6 classes can't have a real constructor. *sigh*

    m_hWndUserControl = hWndUserControl

    FindIEFrame

    If m_hWndIEFrame = -1 Then
        Exit Sub
    End If

    AttachMessage Me, m_hWndIEFrame, WM_EXITSIZEMOVE
    AttachMessage Me, m_hWndIEFrame, WM_SIZE
    AttachMessage Me, m_hWndIEFrame, WM_COMMAND
    AttachMessage Me, m_hWndIEFrame, WM_INITMENUPOPUP

End Sub

Public Sub Destroy()

    DetachMessage Me, m_hWndIEFrame, WM_EXITSIZEMOVE
    DetachMessage Me, m_hWndIEFrame, WM_SIZE
    DetachMessage Me, m_hWndIEFrame, WM_COMMAND
    DetachMessage Me, m_hWndIEFrame, WM_INITMENUPOPUP

End Sub

Private Sub FindIEFrame()

  Dim strClassName As String

    m_hWndInternetExplorerServer = FindInternetExplorerServer()

    If m_hWndInternetExplorerServer = -1 Then
        m_hWndIEFrame = -1
        Exit Sub
    End If

    m_hWndShellDocObjectView = FindShellDocObjectView(m_hWndInternetExplorerServer)

    If m_hWndShellDocObjectView = -1 Then
        m_hWndIEFrame = -1
        Exit Sub
    End If

    m_hWndIEFrame = GetParent(m_hWndShellDocObjectView)

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "IEFrame: GetParent failed in search for IEFrame: " & Hex$(Err.LastDllError)
        m_hWndIEFrame = -1
        Exit Sub
    End If

    strClassName = Space$(128)
    GetClassNameA m_hWndIEFrame, strClassName, 128

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "IEFrame: GetClassName failed in search for IEFrame: " & Hex$(Err.LastDllError)
        m_hWndIEFrame = -1
        Exit Sub
    End If

    strClassName = TrimNull(strClassName)

    If strClassName <> "IEFrame" Then
        modLogging.WriteLineToLog "IEFrame: Failed to find IEFrame. Found instead: " & strClassName
        m_hWndIEFrame = -1
        Exit Sub
    End If

    modLogging.WriteLineToLog "IEFrame: IEFrame hWnd is " & Hex$(m_hWndIEFrame)
    modLogging.WriteLineToLog "IEFrame: Success!"

End Sub

Private Function FindInternetExplorerServer() As Long

  Dim hWnd As Long
  Dim strClassName As String

    If m_hWndUserControl = -1 Then
        modLogging.WriteLineToLog "IEFrame: Don't know the hWnd of the user control."
        FindInternetExplorerServer = -1
        Exit Function
    End If

    hWnd = GetParent(m_hWndUserControl)

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "IEFrame: GetParent failed in search for InternetExplorer_Server: " & Hex$(Err.LastDllError)
        FindInternetExplorerServer = -1
        Exit Function
    End If

    strClassName = Space$(128)
    GetClassNameA hWnd, strClassName, 128

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "IEFrame: GetClassName failed in search for InternetExplorer_Server: " & Hex$(Err.LastDllError)
        FindInternetExplorerServer = -1
        Exit Function
    End If

    strClassName = TrimNull(strClassName)

    If strClassName <> "Internet Explorer_Server" Then
        modLogging.WriteLineToLog "IEFrame: Failed to find Internet Explorer_Server. Found instead: " & strClassName
        FindInternetExplorerServer = -1
        Exit Function
    End If

    modLogging.WriteLineToLog "IEFrame: Internet Explorer_Server hWnd is " & Hex$(hWnd)

    FindInternetExplorerServer = hWnd

End Function

Private Function FindShellDocObjectView(ByVal hWndInternetExplorerServer As Long) As Long

  Dim hWnd As Long
  Dim strClassName As String

    hWnd = GetParent(hWndInternetExplorerServer)

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "IEFrame: GetParent failed in search for Shell DocObject View: " & Hex$(Err.LastDllError)
        FindShellDocObjectView = -1
        Exit Function
    End If

    strClassName = Space$(128)
    GetClassNameA hWnd, strClassName, 128

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "IEFrame: GetClassName failed in search for Shell DocObject View: " & Hex$(Err.LastDllError)
        FindShellDocObjectView = -1
        Exit Function
    End If

    strClassName = TrimNull(strClassName)

    If strClassName <> "Shell DocObject View" Then
        modLogging.WriteLineToLog "IEFrame: Failed to find Shell DocObject View. Found instead: " & strClassName
        FindShellDocObjectView = -1
        Exit Function
    End If

    modLogging.WriteLineToLog "IEFrame: Shell DocObject View hWnd is " & Hex$(hWnd)

    FindShellDocObjectView = hWnd

End Function

Public Property Get hWndIEFrame() As Long

    hWndIEFrame = m_hWndIEFrame

End Property

Public Property Get hWndInternetExplorerServer() As Long

    hWndInternetExplorerServer = m_hWndInternetExplorerServer

End Property

Public Property Get hWndShellDocObjectView() As Long

    hWndShellDocObjectView = m_hWndShellDocObjectView

End Property

Private Property Get ISubclass_MsgResponse() As SSubTUP.EMsgResponse

    ISubclass_MsgResponse = emrConsume

End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTUP.EMsgResponse)

  ' Unused

End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  ' WM_INITMENUPOPUP seems to be the most sane way to get hMenu.
  ' IE6's menu bar is actually a ReBar and the menu buttons are
  ' toolbar buttons that pop up menus at the right coordinates.
  ' IE uses an elaborate procedure to dynamically build the menu
  ' from several menu fragments, depending on whether it's being
  ' used as part of Explorer to show the filesystem or whether
  ' it's being used as a web browser.
  ' ref. https://learn.microsoft.com/en-us/windows/win32/controls/cc-faq-iemenubar

    If iMsg = WM_INITMENUPOPUP Then
        'ISubclass_WindowProc = SSubTUP.CallOldWindowProc(hwnd, iMsg, wParam, lParam)
        ModifyIEMenu wParam
        Exit Function
    End If

    ' The WM_COMMAND processing takes care of redirecting menu
    ' item clicks and keyboard accelerators and forwards them
    ' to events that reach the frontend.

    If iMsg = WM_COMMAND Then
  Dim intHiWord As Integer
  Dim intLoWord As Integer

        intHiWord = HiWord(wParam) ' Menu: 0; Keyboard accelerator: 1
        intLoWord = LoWord(wParam) ' Specifies the identifier

        If intHiWord = 0 Then
            modLogging.WriteLineToLog "IEFrame: Intercept menu WM_COMMAND. ID is:" & Hex$(intLoWord)

  Dim command As IECommand

            command = TranslateToIECommand(intLoWord)

            If command <> IECommand.CommandNone Then
                RaiseEvent CommandReceived(command)
                Exit Function
            End If

            ISubclass_WindowProc = SSubTUP.CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            Exit Function
        End If

        ISubclass_WindowProc = SSubTUP.CallOldWindowProc(hWnd, iMsg, wParam, lParam)
    End If

    ' The idea of the following exercise is to get WindowResized event
    ' to trigger only when the size of the window has actually changed.
    ' That is to say, *only* after the fact. Not before. Not during.

    If iMsg = WM_EXITSIZEMOVE Then
        ISubclass_WindowProc = SSubTUP.CallOldWindowProc(hWnd, iMsg, wParam, lParam)

        If m_GotSizeMessage Then
            ' This way, because WM_EXITSIZEMOVE won't tell you whether
            ' it was a window sizing or window moving operation that
            ' has just finished. *sigh*
            m_GotSizeMessage = False
            RaiseEvent WindowResized
        End If
        Exit Function
    End If

    If iMsg = WM_SIZE Then
        ISubclass_WindowProc = SSubTUP.CallOldWindowProc(hWnd, iMsg, wParam, lParam)

        m_GotSizeMessage = True

        ' Window was minimized
        If wParam = SIZE_MINIMIZED Then
            m_WasMinimized = True
            Exit Function
        End If

        ' Window was not maximized and is now no longer minimized
        If wParam = SIZE_RESTORED And m_WasMinimized Then
            m_WasMinimized = False
            Exit Function
        End If

        ' Window was maximized before and now it's no longer minimized
        If wParam = SIZE_MAXIMIZED And m_WasMinimized Then
            m_WasMinimized = False
            Exit Function
        End If

        ' Window was just maximized
        If wParam = SIZE_MAXIMIZED Then
            m_WasMaximized = True
            RaiseEvent WindowResized
            Exit Function
        End If

        ' Window was just restored from being maximized
        If wParam = SIZE_RESTORED And m_WasMaximized Then
            m_WasMaximized = False
            RaiseEvent WindowResized
            Exit Function
        End If
    End If

End Function

Public Property Get MenuEditCopyEnabled() As Boolean

    MenuEditCopyEnabled = m_MenuEditCopyEnabled

End Property

Public Property Let MenuEditCopyEnabled(enabled As Boolean)

    m_MenuEditCopyEnabled = enabled

End Property

Public Property Get MenuEditCutEnabled() As Boolean

    MenuEditCutEnabled = m_MenuEditCutEnabled

End Property

Public Property Let MenuEditCutEnabled(enabled As Boolean)

    m_MenuEditCutEnabled = enabled

End Property

Public Property Get MenuEditPasteEnabled() As Boolean

    MenuEditPasteEnabled = m_MenuEditPasteEnabled

End Property

Public Property Let MenuEditPasteEnabled(enabled As Boolean)

    m_MenuEditPasteEnabled = enabled

End Property

Public Property Get MenuEditRefreshEnabled() As Boolean

    MenuEditRefreshEnabled = m_MenuEditRefreshEnabled

End Property

Public Property Let MenuEditRefreshEnabled(enabled As Boolean)

    m_MenuEditRefreshEnabled = enabled

End Property

Public Property Get MenuEditStopEnabled() As Boolean

    MenuEditStopEnabled = m_MenuEditStopEnabled

End Property

Public Property Let MenuEditStopEnabled(enabled As Boolean)

    m_MenuEditStopEnabled = enabled

End Property

Private Sub ModifyIEMenu(ByVal hMenu As Long)

  ' These need to be turned off. Stuff like "Java Console" makes no sense.
  ' Other stuff I just have no time for right now. :-(

  ' MENUITEM "&Open...\tCtrl+O", 256, MFT_STRING, MFS_ENABLED
  ' MENUITEM "E&dit", 41030, MFT_STRING, MFS_ENABLED
  ' MENUITEM "&Save\tCtrl+S", 257, MFT_STRING, MFS_GRAYED
  ' MENUITEM "Save &As...", 258, MFT_STRING, MFS_ENABLED
  ' MENUITEM "Page Set&up...", 259, MFT_STRING, MFS_GRAYED
  ' MENUITEM "&Print...\tCtrl+P", 260, MFT_STRING, MFS_GRAYED
  ' MENUITEM "Print Pre&view...", 277, MFT_STRING, MFS_GRAYED
  ' Popup "S&end", 281, MFT_STRING, MFS_ENABLED, 0
  ' MENUITEM "&Work Offline", 40998, MFT_STRING, MFS_ENABLED
  ' Popup "Te&xt Size", 41497, MFT_STRING, MFS_ENABLED, 0
  ' Popup "Enco&ding", 41502, MFT_STRING, MFS_ENABLED, 0
  ' MENUITEM "Pri&vacy Report...", 41507, MFT_STRING, MFS_ENABLED
  ' MENUITEM "&Java Console", 41500, MFT_STRING, MFS_ENABLED

    EnableMenuItem hMenu, 256, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 41030, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 257, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 258, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 259, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 260, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 277, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 281, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 41497, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 40998, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 41502, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 41507, MF_BYCOMMAND Or MF_GRAYED
    EnableMenuItem hMenu, 41500, MF_BYCOMMAND Or MF_GRAYED

    ' Select menu items can be enabled/disabled
    EnableMenuItem hMenu, IE_MENU_COMMAND_EDIT_CUT, IIf(m_MenuEditCutEnabled, MF_BYCOMMAND Or MF_ENABLED, MF_BYCOMMAND Or MF_GRAYED)
    EnableMenuItem hMenu, IE_MENU_COMMAND_EDIT_COPY, IIf(m_MenuEditCopyEnabled, MF_BYCOMMAND Or MF_ENABLED, MF_BYCOMMAND Or MF_GRAYED)
    EnableMenuItem hMenu, IE_MENU_COMMAND_EDIT_PASTE, IIf(m_MenuEditPasteEnabled, MF_BYCOMMAND Or MF_ENABLED, MF_BYCOMMAND Or MF_GRAYED)
    EnableMenuItem hMenu, IE_MENU_COMMAND_EDIT_STOP, IIf(m_MenuEditStopEnabled, MF_BYCOMMAND Or MF_ENABLED, MF_BYCOMMAND Or MF_GRAYED)
    EnableMenuItem hMenu, IE_MENU_COMMAND_EDIT_REFRESH, IIf(m_MenuEditRefreshEnabled, MF_BYCOMMAND Or MF_ENABLED, MF_BYCOMMAND Or MF_GRAYED)

End Sub

Private Function TranslateFromIECommand(command As IECommand) As Long

    Select Case command
      Case IECommand.CommandFileProperties
        TranslateFromIECommand = IE_MENU_COMMAND_FILE_PROPERTIES
      Case IECommand.CommandEditCut
        TranslateFromIECommand = IE_MENU_COMMAND_EDIT_CUT
      Case IECommand.CommandEditCopy
        TranslateFromIECommand = IE_MENU_COMMAND_EDIT_COPY
      Case IECommand.CommandEditPaste
        TranslateFromIECommand = IE_MENU_COMMAND_EDIT_PASTE
      Case IECommand.CommandEditFind
        TranslateFromIECommand = IE_MENU_COMMAND_EDIT_STOP
      Case IECommand.CommandEditStop
        TranslateFromIECommand = IE_MENU_COMMAND_EDIT_REFRESH
      Case IECommand.CommandEditRefresh
        TranslateFromIECommand = IE_MENU_COMMAND_EDIT_FIND
      Case IECommand.CommandViewSource
        TranslateFromIECommand = IE_MENU_COMMAND_VIEW_SOURCE
      Case IECommand.CommandFavoritesAdd
        TranslateFromIECommand = IE_MENU_COMMAND_FAVORITES_ADD
      Case Else
        TranslateFromIECommand = -1
    End Select

End Function

Private Function TranslateToIECommand(LoWord As Integer) As IECommand

    Select Case LoWord
      Case IE_MENU_COMMAND_FILE_PROPERTIES
        TranslateToIECommand = IECommand.CommandFileProperties
      Case IE_MENU_COMMAND_EDIT_CUT
        TranslateToIECommand = IECommand.CommandEditCut
      Case IE_MENU_COMMAND_EDIT_COPY
        TranslateToIECommand = IECommand.CommandEditCopy
      Case IE_MENU_COMMAND_EDIT_PASTE
        TranslateToIECommand = IECommand.CommandEditPaste
      Case IE_MENU_COMMAND_EDIT_STOP
        TranslateToIECommand = IECommand.CommandEditStop
      Case IE_MENU_COMMAND_EDIT_REFRESH
        TranslateToIECommand = IECommand.CommandEditRefresh
      Case IE_MENU_COMMAND_EDIT_FIND
        TranslateToIECommand = IECommand.CommandEditFind
      Case IE_MENU_COMMAND_VIEW_SOURCE
        TranslateToIECommand = IECommand.CommandViewSource
      Case IE_MENU_COMMAND_FAVORITES_ADD
        TranslateToIECommand = IECommand.CommandFavoritesAdd
      Case Else
        TranslateToIECommand = IECommand.CommandNone
    End Select

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-12 08:13)  Decl: 47  Code: 447  Total: 494 Lines
':) CommentOnly: 43 (8.7%)  Commented: 5 (1%)  Filled: 361 (73.1%)  Empty: 133 (26.9%)  Max Logic Depth: 4
