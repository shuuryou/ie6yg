VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEAddressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_hWndIEFrame As Long
Private m_hWndAddressBar As Long

Public Sub Constructor(hWndIEFrame As Long)
    ' VB6 classes can't have a real constructor. *sigh*
    
    m_hWndIEFrame = hWndIEFrame

    FindAddressBar
End Sub

Public Property Get hWnd() As Long
    hWnd = m_hWndAddressBar
End Property

Private Sub FindAddressBar()
  Dim hWnd As Long

    If m_hWndIEFrame = -1 Then
        modLogging.WriteLineToLog "FindAddressBar: m_hWndIEFrame is bad."
        Exit Sub      '---> Bottom
    End If

    ' ------------------------------------------------------------------------

    ' Now start going down again and find the first "WorkerW"

    hWnd = FindWindowEx(m_hWndIEFrame, 0&, "WorkerW", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindAddressBar: FindWindowEx WorkerW failed. GetLastError returns " & Hex$(Err.LastDllError)
        m_hWndAddressBar = -1
        Exit Sub      '---> Bottom
    End If

    modLogging.WriteLineToLog "FindAddressBar: WorkerW hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the "ReBarWindow32" containing the address bar

    hWnd = FindWindowEx(hWnd, 0&, "ReBarWindow32", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindAddressBar: FindWindowEx ReBarWindow32 failed. GetLastError returns " & Hex$(Err.LastDllError)
        m_hWndAddressBar = -1
        Exit Sub      '---> Bottom
    End If

    modLogging.WriteLineToLog "FindAddressBar: ReBarWindow32 hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the "ComboBoxEx32" to get to the address bar

    hWnd = FindWindowEx(hWnd, 0&, "ComboBoxEx32", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindAddressBar: FindWindowEx ComboBoxEx32 failed. GetLastError returns " & Hex$(Err.LastDllError)
        m_hWndAddressBar = -1
        Exit Sub      '---> Bottom
    End If

    modLogging.WriteLineToLog "FindAddressBar: ComboBoxEx32 hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the actual "ComboBox" to get to the edit control

    hWnd = FindWindowEx(hWnd, 0&, "ComboBox", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindAddressBar: FindWindowEx ComboBox failed. GetLastError returns " & Hex$(Err.LastDllError)
        m_hWndAddressBar = -1
        Exit Sub      '---> Bottom
    End If

    modLogging.WriteLineToLog "FindAddressBar: ComboBox hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Finally find the "Edit" window and we're there

    hWnd = FindWindowEx(hWnd, 0&, "Edit", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindAddressBar: FindWindowEx Edit failed. GetLastError returns " & Hex$(Err.LastDllError)
        m_hWndAddressBar = -1
        Exit Sub      '---> Bottom
    End If

    modLogging.WriteLineToLog "FindAddressBar: Edit hWnd is " & Hex$(hWnd)

    m_hWndAddressBar = hWnd
End Sub

Public Sub SetText(newText As String)

    If m_hWndAddressBar = -1 Then
        modLogging.WriteLineToLog "SetAddressBarText: Cannot set if address bar not found."
        Exit Sub '---> Bottom
    End If

    modLogging.WriteLineToLog "SetAddressBarText: Text becomes: " & newText

    SendMessage m_hWndAddressBar, WM_SETTEXT, 0&, ByVal newText
End Sub
