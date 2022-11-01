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

Implements SSubTimer.ISubclass

Private m_hWndUserControl As Long
Private m_hWndInternetExplorerServer As Long
Private m_hWndShellDocObjectView As Long
Private m_hWndIEFrame As Long

Public Event WindowResizing()
Public Event WindowResized()

Private Sub Class_Terminate()

    If m_hWndIEFrame = -1 Then
        Exit Sub
    End If

    DetachMessage Me, m_hWndIEFrame, WM_SIZE
    DetachMessage Me, m_hWndIEFrame, WM_EXITSIZEMOVE

End Sub

Public Sub Constructor(hWndUserControl As Long)

  ' VB6 classes can't have a real constructor. *sigh*

    m_hWndUserControl = hWndUserControl

    FindIEFrame

    If m_hWndIEFrame = -1 Then
        Exit Sub
    End If

    AttachMessage Me, m_hWndIEFrame, WM_SIZE
    AttachMessage Me, m_hWndIEFrame, WM_EXITSIZEMOVE

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
        modLogging.WriteLineToLog "FindAndSubclassIEFrame: GetParent failed: " & Hex$(Err.LastDllError)
        m_hWndIEFrame = -1
        Exit Sub
    End If

    strClassName = Space$(128)
    GetClassNameA m_hWndIEFrame, strClassName, 128

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "FindAndSubclassIEFrame: GetClassName failed: " & Hex$(Err.LastDllError)
        m_hWndIEFrame = -1
        Exit Sub
    End If

    strClassName = TrimNull(strClassName)

    If strClassName <> "IEFrame" Then
        modLogging.WriteLineToLog "FindAndSubclassIEFrame: Failed (" & strClassName & ")"
        m_hWndIEFrame = -1
        Exit Sub
    End If

    modLogging.WriteLineToLog "FindAndSubclassIEFrame: IEFrame hWnd is " & Hex$(m_hWndIEFrame)

End Sub

Private Function FindInternetExplorerServer() As Long

  Dim hWnd As Long
  Dim strClassName As String

    If m_hWndUserControl = -1 Then
        modLogging.WriteLineToLog "FindInternetExplorerServer: Don't know the hWnd of the user control yet."
        FindInternetExplorerServer = -1
        Exit Function
    End If

    hWnd = GetParent(m_hWndUserControl)

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "FindInternetExplorerServer: GetParent failed: " & Hex$(Err.LastDllError)
        FindInternetExplorerServer = -1
        Exit Function
    End If

    strClassName = Space$(128)
    GetClassNameA hWnd, strClassName, 128

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "FindInternetExplorerServer: GetClassName failed: " & Hex$(Err.LastDllError)
        FindInternetExplorerServer = -1
        Exit Function
    End If

    strClassName = TrimNull(strClassName)

    If strClassName <> "Internet Explorer_Server" Then
        modLogging.WriteLineToLog "FindInternetExplorerServer: Failed (" & strClassName & ")"
        FindInternetExplorerServer = -1
        Exit Function
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
        Exit Function
    End If

    strClassName = Space$(128)
    GetClassNameA hWnd, strClassName, 128

    If Err.LastDllError <> 0 Then
        modLogging.WriteLineToLog "FindShellDocObjectView: GetClassName failed: returns " & Hex$(Err.LastDllError)
        FindShellDocObjectView = -1
        Exit Function
    End If

    strClassName = TrimNull(strClassName)

    If strClassName <> "Shell DocObject View" Then
        modLogging.WriteLineToLog "FindShellDocObjectView: Failed (" & strClassName & ")"
        FindShellDocObjectView = -1
        Exit Function
    End If

    modLogging.WriteLineToLog "FindShellDocObjectView: Shell DocObject View hWnd is " & Hex$(hWnd)

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

Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse

    ISubclass_MsgResponse = emrConsume

End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)

  ' Unused

End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal WParam As Long, ByVal lParam As Long) As Long

  ' There is a bug here that required a patched version of SSubTimer
  ' to resolve completely. Read more details about it in this file:
  ' ..\Support\SSubTimerUnicode\READ.TXT

    If iMsg = WM_SIZE Then
        RaiseEvent WindowResizing
      ElseIf iMsg = WM_EXITSIZEMOVE Then
        RaiseEvent WindowResized
    End If

    CallOldWindowProc hWnd, iMsg, WParam, lParam

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-01 16:03)  Decl: 11  Code: 202  Total: 213 Lines
':) CommentOnly: 7 (3,3%)  Commented: 0 (0%)  Filled: 152 (71,4%)  Empty: 61 (28,6%)  Max Logic Depth: 2
