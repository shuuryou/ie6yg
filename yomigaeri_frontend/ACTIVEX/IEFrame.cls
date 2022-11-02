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

Public Event WindowResized()

Public Sub Construct(hWndUserControl As Long)

  ' VB6 classes can't have a real constructor. *sigh*

    m_hWndUserControl = hWndUserControl

    FindIEFrame

    If m_hWndIEFrame = -1 Then
        Exit Sub
    End If

    AttachMessage Me, m_hWndIEFrame, WM_EXITSIZEMOVE
    AttachMessage Me, m_hWndIEFrame, WM_SIZE

End Sub

Public Sub Destroy()

    DetachMessage Me, m_hWndIEFrame, WM_EXITSIZEMOVE
    DetachMessage Me, m_hWndIEFrame, WM_SIZE

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

    ISubclass_MsgResponse = emrPreprocess

End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTUP.EMsgResponse)

  ' Unused

End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  ' The idea of the following exercise is to get WindowResized event
  ' to trigger only when the size of the window has actually changed.
  ' That is to say, after the fact. Not before. Not during.

    If iMsg = WM_EXITSIZEMOVE Then
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

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-03 00:39)  Decl: 14  Code: 238  Total: 252 Lines
':) CommentOnly: 15 (6%)  Commented: 0 (0%)  Filled: 176 (69,8%)  Empty: 76 (30,2%)  Max Logic Depth: 3
