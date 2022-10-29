VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEBrowserManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function AccessibleObjectFromWindow Lib "oleacc.dll" (ByVal hWnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal ndx As Long, ByVal newValue As Long) As Long

Private Type TBBUTTON
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    dwData As Long
    iString As Long
End Type

Private Const IIDSTR_IHTMLElement As String = "{3050f1ff-98b5-11cf-bb82-00aa00bdce0b}"
Private Const IIDSTR_IWebBrowser2 As String = "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"
Private Const IIDSTR_SWebBrowserApp As String = "{0002DF05-0000-0000-C000-000000000046}"

Private Const OBJID_CLIENT As Long = -4

Private Const GWL_WNDPROC As Long = -4

Private Const WM_SETTEXT As Long = &HC

Private Const WM_USER As Long = &H400
Private Const TB_BUTTONCOUNT As Long = (WM_USER + &H18)
Private Const TB_GETBUTTON As Long = (WM_USER + &H17)
Private Const TB_ENABLEBUTTON  As Long = (WM_USER + &H1)

' These are the TBBUTTON.iBitmap values, at least in IE6. I don't know if it's
' a good idea to use iBitmap values, but the values of idCommand will change
' randomly when toolbar buttons are added, removed, or moved around. Toolbar
' bitmap indexes appear to always remain stable.
Private Const TOOLBAR_IBITMAP_BACK As Byte = 0
Private Const TOOLBAR_IBITMAP_FORWARD As Byte = 1
Private Const TOOLBAR_IBITMAP_STOP As Byte = 2
Private Const TOOLBAR_IBITMAP_REFRESH As Byte = 3
Private Const TOOLBAR_IBITMAP_HOME As Byte = 4
'Private Const TOOLBAR_IBITMAP_SEARCH As Byte = 5
'Private Const TOOLBAR_IBITMAP_FAVORITES As Byte = 6
'Private Const TOOLBAR_IBITMAP_PRINT As Byte = 7
'Private Const TOOLBAR_IBITMAP_MEDIA As Byte = 11
'Private Const TOOLBAR_IBITMAP_HISTORY As Byte = 12

Private TOOLBAR_IDCOMMAND_BACK As Integer
Private TOOLBAR_IDCOMMAND_FORWARD As Integer
Private TOOLBAR_IDCOMMAND_STOP As Integer
Private TOOLBAR_IDCOMMAND_REFRESH As Integer

Private TOOLBAR_BUTTON_STATE_BACK As Boolean
Private TOOLBAR_BUTTON_STATE_FORWARD As Boolean
Private TOOLBAR_BUTTON_STATE_STOP As Boolean
Private TOOLBAR_BUTTON_STATE_REFRESH As Boolean

Private m_History As IUrlHistoryStg
Private WithEvents m_Browser As WebBrowser
Attribute m_Browser.VB_VarHelpID = -1

Private m_hWndUserControl As Long
Private m_hWndAddressBar As Long
Private m_hWndButtonToolbar As Long
Private m_hWndReBarWindow32 As Long
Private m_hOldWndProcButtonToolbar As Long
Private m_hOldWndProcReBarWindow32 As Long

Public Event IEWantsToNavigate(newUrl As String)
Public Event IEToolbarCommandClicked(commandId As Long)

Private Sub Class_Initialize()

    m_hWndUserControl = -1
    m_hWndAddressBar = -1
    m_hWndButtonToolbar = -1
    m_hWndReBarWindow32 = -1
    m_hOldWndProcButtonToolbar = -1
    m_hOldWndProcReBarWindow32 = -1

End Sub

Public Property Let hWndUserControl(ByRef hWnd As Long)

    m_hWndUserControl = hWnd

End Property

Public Property Get hWndUserControl() As Long

    hWndUserControl = m_hWndUserControl

End Property

Public Property Let ToolbarButtonStateBack(state As Boolean)

    TOOLBAR_BUTTON_STATE_BACK = state

    If m_hWndButtonToolbar <> -1 Then
        SendMessage m_hWndButtonToolbar, TB_ENABLEBUTTON, TOOLBAR_IDCOMMAND_BACK, Abs(state)
    End If

End Property

Public Property Get ToolbarButtonStateBack() As Boolean

    ToolbarButtonStateBack = TOOLBAR_BUTTON_STATE_BACK

End Property

Public Property Let ToolbarButtonStateForward(state As Boolean)

    TOOLBAR_BUTTON_STATE_FORWARD = state

    If m_hWndButtonToolbar <> -1 Then
        SendMessage m_hWndButtonToolbar, TB_ENABLEBUTTON, TOOLBAR_IDCOMMAND_FORWARD, Abs(state)
    End If

End Property

Public Property Get ToolbarButtonStateForward() As Boolean

    ToolbarButtonStateForward = TOOLBAR_BUTTON_STATE_FORWARD

End Property

Public Property Let ToolbarButtonStateStop(state As Boolean)

    TOOLBAR_BUTTON_STATE_STOP = state

    If m_hWndButtonToolbar <> -1 Then
        SendMessage m_hWndButtonToolbar, TB_ENABLEBUTTON, TOOLBAR_IDCOMMAND_STOP, Abs(state)
    End If

End Property

Public Property Get ToolbarButtonStateStop() As Boolean

    ToolbarButtonStateStop = TOOLBAR_BUTTON_STATE_STOP

End Property

Public Property Let ToolbarButtonStateRefresh(state As Boolean)

    TOOLBAR_BUTTON_STATE_REFRESH = state

    If m_hWndButtonToolbar <> -1 Then
        SendMessage m_hWndButtonToolbar, TB_ENABLEBUTTON, TOOLBAR_IDCOMMAND_REFRESH, Abs(state)
    End If

End Property

Public Property Get ToolbarButtonStateRefresh() As Boolean

    ToolbarButtonStateRefresh = TOOLBAR_BUTTON_STATE_REFRESH

End Property

Public Property Get ToolbarIdCommandBack() As Long

    ToolbarIdCommandBack = TOOLBAR_IDCOMMAND_BACK

End Property

Public Property Get ToolbarIdCommandForward() As Long

    ToolbarIdCommandForward = TOOLBAR_IDCOMMAND_FORWARD

End Property

Public Property Get ToolbarIdCommandStop() As Long

    ToolbarIdCommandStop = TOOLBAR_IDCOMMAND_STOP

End Property

Public Property Get ToolbarIdCommandRefresh() As Long

    ToolbarIdCommandRefresh = TOOLBAR_IDCOMMAND_REFRESH

End Property

Public Property Get ToolbarOldWndProc() As Long

    ToolbarOldWndProc = m_hOldWndProcButtonToolbar

End Property

Public Property Get ReBarOldWndProc()

    ReBarOldWndProc = m_hOldWndProcReBarWindow32

End Property

Public Property Get RebarWindow32OldWndProc() As Long

    RebarWindow32OldWndProc = m_hOldWndProcReBarWindow32

End Property

Private Function FindInternetExplorerServer() As Long

  Dim hWnd As Long
  Dim strClassName As String

    hWnd = GetParent(m_hWndUserControl)

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "FindInternetExplorerServer: GetParent failed: " & Hex$(Err.LastDllError)
        FindInternetExplorerServer = -1
        Exit Function '---> Bottom
    End If

    strClassName = Space$(128)
    GetClassName hWnd, strClassName, 128

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "FindInternetExplorerServer: GetClassName failed: " & Hex$(Err.LastDllError)
        FindInternetExplorerServer = -1
        Exit Function '---> Bottom
    End If

    strClassName = TrimNull(strClassName)

    If strClassName <> "Internet Explorer_Server" Then
        modLogging.WriteLineToLog "FindInternetExplorerServer: Failed (" & strClassName & ")"
        FindInternetExplorerServer = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "FindInternetExplorerServer: Internet Explorer_Server hWnd is " & Hex$(hWnd)

    FindInternetExplorerServer = hWnd

End Function

Private Function FindShellDocObjectView(ByVal hWndInternetExplorerServer As Long) As Long

  Dim hWnd As Long
  Dim strClassName As String

    hWnd = GetParent(hWndInternetExplorerServer)

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "FindShellDocObjectView: GetParent failed: " & Hex$(Err.LastDllError)
        FindShellDocObjectView = -1
        Exit Function '---> Bottom
    End If

    strClassName = Space$(128)
    GetClassName hWnd, strClassName, 128

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "FindShellDocObjectView: GetClassName failed: returns " & Hex$(Err.LastDllError)
        FindShellDocObjectView = -1
        Exit Function '---> Bottom
    End If

    strClassName = TrimNull(strClassName)

    If strClassName <> "Shell DocObject View" Then
        modLogging.WriteLineToLog "FindShellDocObjectView: Failed (" & strClassName & ")"
        FindShellDocObjectView = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "FindShellDocObjectView: Shell DocObject View hWnd is " & Hex$(hWnd)

    FindShellDocObjectView = hWnd

End Function

Private Function FindIEFrame() As Long

  Dim hWnd As Long
  Dim strClassName As String

    hWnd = FindInternetExplorerServer()

    If hWnd = -1 Then
        Exit Function '---> Bottom
    End If

    hWnd = FindShellDocObjectView(hWnd)

    If hWnd = -1 Then
        Exit Function '---> Bottom
    End If

    hWnd = GetParent(hWnd)

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "FindIEFrame: GetParent failed: " & Hex$(Err.LastDllError)
        FindIEFrame = -1
        Exit Function '---> Bottom
    End If

    strClassName = Space$(128)
    GetClassName hWnd, strClassName, 128

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "FindIEFrame: GetClassName failed: " & Hex$(Err.LastDllError)
        FindIEFrame = -1
        Exit Function '---> Bottom
    End If

    strClassName = TrimNull(strClassName)

    If strClassName <> "IEFrame" Then
        modLogging.WriteLineToLog "FindIEFrame: Failed (" & strClassName & ")"
        FindIEFrame = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "FindIEFrame: IEFrame hWnd is " & Hex$(hWnd)

    FindIEFrame = hWnd

End Function

Private Function FindAddressBar() As Long

  Dim hWnd As Long

    hWnd = FindIEFrame()

    If hWnd = -1 Then
        Exit Function '---> Bottom
    End If

    ' ------------------------------------------------------------------------

    ' Now start going down again and find the first "WorkerW"

    hWnd = FindWindowEx(hWnd, 0&, "WorkerW", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindAddressBar: FindWindowEx WorkerW failed. GetLastError returns " & Hex$(Err.LastDllError)
        FindAddressBar = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "FindAddressBar: WorkerW hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the "ReBarWindow32" containing the address bar

    hWnd = FindWindowEx(hWnd, 0&, "ReBarWindow32", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindAddressBar: FindWindowEx ReBarWindow32 failed. GetLastError returns " & Hex$(Err.LastDllError)
        FindAddressBar = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "FindAddressBar: ReBarWindow32 hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the "ComboBoxEx32" to get to the address bar

    hWnd = FindWindowEx(hWnd, 0&, "ComboBoxEx32", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindAddressBar: FindWindowEx ComboBoxEx32 failed. GetLastError returns " & Hex$(Err.LastDllError)
        FindAddressBar = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "FindAddressBar: ComboBoxEx32 hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the actual "ComboBox" to get to the edit control

    hWnd = FindWindowEx(hWnd, 0&, "ComboBox", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindAddressBar: FindWindowEx ComboBox failed. GetLastError returns " & Hex$(Err.LastDllError)
        FindAddressBar = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "FindAddressBar: ComboBox hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Finally find the "Edit" window and we're there

    hWnd = FindWindowEx(hWnd, 0&, "Edit", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindAddressBar: FindWindowEx Edit failed. GetLastError returns " & Hex$(Err.LastDllError)
        FindAddressBar = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "FindAddressBar: Edit hWnd is " & Hex$(hWnd)

    FindAddressBar = hWnd

End Function

Public Sub HookButtonToolbarCommands()

  Dim hwndIEFrame As Long

  Dim hWndTemp As Long
  Dim hwndToolbar As Long

  Dim sctTbButton As TBBUTTON
  Dim intButtonCount As Integer

  Dim i As Integer, j As Integer

    hwndIEFrame = FindIEFrame()
    If hwndIEFrame = -1 Then
        m_hWndButtonToolbar = -1
        Exit Sub '---> Bottom
    End If

    hWndTemp = FindWindowEx(hwndIEFrame, 0&, "WorkerW", vbNullString)

    If hWndTemp = 0 Then
        modLogging.WriteLineToLog "HookButtonToolbarCommands: WorkerW not found."
        m_hWndButtonToolbar = -1
        Exit Sub '---> Bottom
    End If

    hWndTemp = FindWindowEx(hWndTemp, 0&, "ReBarWindow32", vbNullString)

    If hWndTemp = 0 Then
        modLogging.WriteLineToLog "HookButtonToolbarCommands: ReBarWindow32 not found."
        m_hWndButtonToolbar = -1
        Exit Sub '---> Bottom
    End If

    modLogging.WriteLineToLog "HookButtonToolbarCommands: Found ReBarWindow32: " & Hex$(hWndTemp)
    m_hWndReBarWindow32 = hWndTemp

    hwndToolbar = 0&

NextToolbar:
    TOOLBAR_IDCOMMAND_BACK = -1
    TOOLBAR_IDCOMMAND_FORWARD = -1
    TOOLBAR_IDCOMMAND_STOP = -1
    TOOLBAR_IDCOMMAND_REFRESH = -1

    hwndToolbar = FindWindowEx(hWndTemp, hwndToolbar, "ToolbarWindow32", vbNullString)

    If hwndToolbar = 0 Then
        modLogging.WriteLineToLog "HookButtonToolbarCommands: ToolbarWindow32 not found or all of them were wrong."
        m_hWndButtonToolbar = -1
        m_hWndReBarWindow32 = -1
        Exit Sub '---> Bottom
    End If

    intButtonCount = SendMessage(hwndToolbar, TB_BUTTONCOUNT, 0, 0)

    For i = 0 To intButtonCount - 1 Step 1
        SendMessage hwndToolbar, TB_GETBUTTON, i, sctTbButton

        ' Another shitty hack. It would have been nice to use TB_GETBUTTONTEXT,
        ' but this consistenly crashed the BHO and IE unless there was a MsgBox
        ' in this loop to interrupt it every iteration. I guess it's a bug in
        ' COMCTL32?

        ' Don't know. Don't care. All of what is going on inside this file is
        ' insane anyway, so I gave up after a day of trying to debug the crash.
        ' BHO's have a very large iBitmap value, so the following is reliable
        ' enough to find the button toolbar.

        ' It doesn't help that most BHO development went down 15-20 years ago
        ' and the knowledge has since become somewhat lost on the Internet.
        If sctTbButton.idCommand <> 0 And sctTbButton.iBitmap > 128 Then
            m_hWndButtonToolbar = hwndToolbar
            modLogging.WriteLineToLog "HookButtonToolbarCommands: Found it: " & Hex$(hwndToolbar)

            ' Now go through the buttons a second time to find the interesting
            ' ones by their bitmap index. The bitmap index seems to be stable.
            ' The command IDs change randomly when changing toolbar settings.

            For j = 0 To intButtonCount - 1 Step 1
                SendMessage hwndToolbar, TB_GETBUTTON, j, sctTbButton

                Select Case sctTbButton.iBitmap
                  Case TOOLBAR_IBITMAP_BACK
                    modLogging.WriteLineToLog "HookButtonToolbarCommands: Found BACK button. Command ID: " & sctTbButton.idCommand
                    TOOLBAR_IDCOMMAND_BACK = sctTbButton.idCommand
                  Case TOOLBAR_IBITMAP_FORWARD
                    modLogging.WriteLineToLog "HookButtonToolbarCommands: Found FORWARD button. Command ID: " & sctTbButton.idCommand
                    TOOLBAR_IDCOMMAND_FORWARD = sctTbButton.idCommand
                  Case TOOLBAR_IBITMAP_STOP
                    modLogging.WriteLineToLog "HookButtonToolbarCommands: Found STOP button. Command ID: " & sctTbButton.idCommand
                    TOOLBAR_IDCOMMAND_STOP = sctTbButton.idCommand
                  Case TOOLBAR_IBITMAP_REFRESH
                    modLogging.WriteLineToLog "HookButtonToolbarCommands: Found REFRESH button. Command ID: " & sctTbButton.idCommand
                    TOOLBAR_IDCOMMAND_REFRESH = sctTbButton.idCommand
                End Select
            Next j

            GoTo DoHook
        End If
    Next i
    
    GoTo NextToolbar

DoHook:
    m_hOldWndProcButtonToolbar = SetWindowLong(m_hWndButtonToolbar, GWL_WNDPROC, AddressOf modWndProc.ButtonToolbarWndProc)
    modLogging.WriteLineToLog "HookButtonToolbarCommands: Hooked ToolbarWindow32 WndProc. Old one was at " & Hex$(m_hOldWndProcButtonToolbar)

    m_hOldWndProcReBarWindow32 = SetWindowLong(m_hWndReBarWindow32, GWL_WNDPROC, AddressOf modWndProc.RebarWindow32WndProc)
    modLogging.WriteLineToLog "HookButtonToolbarCommands: Hooked ReBarWindow32 WndProc. Old one was at " & Hex$(m_hOldWndProcReBarWindow32)

End Sub

Public Sub ReleaseButtonToolbarCommands()

    If m_hOldWndProcButtonToolbar = -1 Then
        Exit Sub '---> Bottom
    End If

    SetWindowLong m_hWndButtonToolbar, GWL_WNDPROC, m_hOldWndProcButtonToolbar
    m_hOldWndProcButtonToolbar = -1
    modLogging.WriteLineToLog "Restored ToolbarWindow32 WndProc."

    SetWindowLong m_hWndReBarWindow32, GWL_WNDPROC, m_hOldWndProcReBarWindow32
    m_hOldWndProcButtonToolbar = -1
    modLogging.WriteLineToLog "Restored ReBarWindow32 WndProc."

End Sub

Public Sub RaiseButtonToolbarEvent(commandId As Long)

    modLogging.WriteLineToLog "RaiseButtonToolbarEvent: Catch and redirect command ID " & CStr(commandId)
    RaiseEvent IEToolbarCommandClicked(commandId)

End Sub

Public Sub SetAddressBarText(newText As String)

    If m_hWndAddressBar = 0 Then
        m_hWndAddressBar = FindAddressBar()
    End If

    If m_hWndAddressBar = -1 Then
        modLogging.WriteLineToLog "SetAddressBarText: Cannot set if address bar not found."
        Exit Sub '---> Bottom
    End If

    modLogging.WriteLineToLog "SetAddressBarText: Text becomes: " & newText

    SendMessage m_hWndAddressBar, WM_SETTEXT, 0&, ByVal newText

End Sub

Public Sub HookIWebBrowser2()

  Dim iid_IDispatch As UUID
  Dim iid_SWebBrowserApp As UUID
  Dim iid_IHTMLElement As UUID
  Dim iid_IWebBrowser2 As UUID

  Dim hWndInternetExplorerServer As Long
  Dim objIEBrowser As Object
  Dim objIServiceProvider As IServiceProvider
  Dim objIServiceProvider2 As IServiceProvider
  Dim objHTMLElement As Object
  Dim objWebBrowser As Object

  ' Convert the GUID strings for the interfaces to interface identifiers

    IIDFromString StrPtr(IIDSTR_IDispatch), iid_IDispatch
    IIDFromString StrPtr(IIDSTR_SWebBrowserApp), iid_SWebBrowserApp
    IIDFromString StrPtr(IIDSTR_IHTMLElement), iid_IHTMLElement
    IIDFromString StrPtr(IIDSTR_IWebBrowser2), iid_IWebBrowser2

    ' We start from the Internet Explorer_Server window
    hWndInternetExplorerServer = FindInternetExplorerServer()

    If hWndInternetExplorerServer = -1 Then
        modLogging.WriteLineToLog "HookIWebBrowser2: hWnd of Internet Explorer_Server not found."
        Exit Sub '---> Bottom
    End If

    ' Documentation for what follows (sort of):
    ' http://jeffcode.blogspot.com/2008/06/internet-explorer-object-cookbook.html
    ' https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/249232
    ' https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/257717

    ' The approach below does it using AccessibleObjectFromWindow (Win98+ required)

    If AccessibleObjectFromWindow(hWndInternetExplorerServer, OBJID_CLIENT, iid_IDispatch, objIEBrowser) <> 0 Then
        modLogging.WriteLineToLog "HookIWebBrowser2: AccessibleObjectFromWindow failed."
        Exit Sub '---> Bottom
    End If

    Set objIServiceProvider = objIEBrowser.accChild(CVar(0&))

    objIServiceProvider.QueryService iid_IHTMLElement, iid_IDispatch, objHTMLElement

    Set objIServiceProvider2 = objHTMLElement.Document

    objIServiceProvider2.QueryService iid_SWebBrowserApp, iid_IWebBrowser2, objWebBrowser

    Set m_Browser = objWebBrowser

    Set objIServiceProvider2 = Nothing
    Set objIServiceProvider = Nothing

    modLogging.WriteLineToLog "HookIWebBrowser2: Success!"

End Sub

Public Sub ReleaseIWebBrowser2()

    Set m_Browser = Nothing

End Sub

Public Sub SetStatusBarText(text As String)

  ' This is unfortunately not reliable very early during ActiveX control
  ' initialization when the page has not yet settled. Therefore it is
  ' done using a JScript event instead.

    If m_Browser Is Nothing Then
        modLogging.WriteLineToLog "SetStatusBarText: IWebBrowser2 not locked onto."
        Exit Sub '---> Bottom
    End If

    m_Browser.StatusText = text
    modLogging.WriteLineToLog "SetStatusBarText: Text becomes: " & text

End Sub

Private Sub m_Browser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    Cancel = True
    modLogging.WriteLineToLog "BeforeNavigate2: Catch and redirect navigating to: " & CStr(URL)
    RaiseEvent IEWantsToNavigate(CStr(URL))

End Sub

Private Sub m_Browser_DownloadBegin()

    m_Browser.Stop

End Sub

Public Sub PushIntoHistory(ByVal URL As String, ByVal title As String)

    If m_History Is Nothing Then
        Set m_History = New UrlHistory
    End If

    modLogging.WriteLineToLog "PushIntoHistory: Add to history URL=" & URL & "; Title=" & title

    m_History.ADDURL URL, title, STATURLFLAG_ISTOPLEVEL

End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Oct-29 22:14)  Decl: 71  Code: 595  Total: 666 Lines
':) CommentOnly: 44 (6.6%)  Commented: 0 (0%)  Filled: 449 (67.4%)  Empty: 217 (32.6%)  Max Logic Depth: 5
