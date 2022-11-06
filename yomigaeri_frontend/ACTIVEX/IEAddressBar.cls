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

Public Sub Construct(hWndIEFrame As Long)

  ' VB6 classes can't have a real constructor. *sigh*

    m_hWndIEFrame = hWndIEFrame

    FindAddressBar

End Sub

Private Sub FindAddressBar()

  Dim hWnd As Long

    If m_hWndIEFrame = -1 Then
        modLogging.WriteLineToLog "IEAddressBar: m_hWndIEFrame is bad."
        Exit Sub
    End If

    ' ------------------------------------------------------------------------

    ' Now start going down again and find the first "WorkerA"

    hWnd = FindWindowEx(m_hWndIEFrame, 0&, "WorkerA", vbNullString)

    If hWnd = 0 Then
        modLogging.WriteLineToLog "IEAddressBar: FindWindowEx WorkerA failed."

        hWnd = FindWindowEx(m_hWndIEFrame, 0&, "WorkerW", vbNullString)

        If hWnd = 0 Then
            modLogging.WriteLineToLog "IEAddressBar: FindWindowEx WorkerW also failed."
            m_hWndAddressBar = -1
            Exit Sub
          Else
            modLogging.WriteLineToLog "IEAddressBar: FindWindowEx WorkerW worked."
        End If
    End If

    modLogging.WriteLineToLog "FindAddressBar: WorkerA hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the "ReBarWindow32" containing the address bar

    hWnd = FindWindowEx(hWnd, 0&, "ReBarWindow32", vbNullString)

    If hWnd = 0 Then
        modLogging.WriteLineToLog "IEAddressBar: FindWindowEx ReBarWindow32 failed."
        m_hWndAddressBar = -1
        Exit Sub
    End If

    modLogging.WriteLineToLog "IEAddressBar: ReBarWindow32 hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the "ComboBoxEx32" to get to the address bar

    hWnd = FindWindowEx(hWnd, 0&, "ComboBoxEx32", vbNullString)

    If hWnd = 0 Then
        modLogging.WriteLineToLog "IEAddressBar: FindWindowEx ComboBoxEx32 failed."
        m_hWndAddressBar = -1
        Exit Sub
    End If

    modLogging.WriteLineToLog "IEAddressBar: ComboBoxEx32 hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the actual "ComboBox" to get to the edit control

    hWnd = FindWindowEx(hWnd, 0&, "ComboBox", vbNullString)

    If hWnd = 0 Then
        modLogging.WriteLineToLog "IEAddressBar: FindWindowEx ComboBox failed."
        m_hWndAddressBar = -1
        Exit Sub
    End If

    modLogging.WriteLineToLog "IEAddressBar: ComboBox hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Finally find the "Edit" window and we're there

    hWnd = FindWindowEx(hWnd, 0&, "Edit", vbNullString)

    If hWnd = 0 Then
        modLogging.WriteLineToLog "IEAddressBar: FindWindowEx Edit failed."
        m_hWndAddressBar = -1
        Exit Sub
    End If

    modLogging.WriteLineToLog "IEAddressBar: Edit hWnd is " & Hex$(hWnd)

    m_hWndAddressBar = hWnd
    
    modLogging.WriteLineToLog "IEAddressBar: Success!"

End Sub

Public Property Get hWnd() As Long

    hWnd = m_hWndAddressBar

End Property

Public Property Get Text() As String

  Dim lngLength As Long
  Dim bytText() As Byte

    If m_hWndAddressBar = -1 Then
        Text = ""
        Exit Property
    End If

    lngLength = SendMessage(m_hWndAddressBar, WM_GETTEXTLENGTH, 0&, 0&)

    If lngLength = 0 Then
        Text = ""
        Exit Property
    End If

    ReDim bytText(0 To lngLength + 1) As Byte

    SendMessage m_hWndAddressBar, WM_GETTEXT, ByVal (lngLength + 1), bytText(0)

    Text = TrimNull(StrConv(bytText, vbUnicode))
    
    modLogging.WriteLineToLog "IEAddressBar: Get text result: " & Text

End Property

Public Property Let Text(newText As String)

  Dim Text() As Byte

    If m_hWndAddressBar = -1 Then
        Exit Property
    End If

    Text = StrConv(newText, vbFromUnicode) & Chr$(0)

    SendMessage m_hWndAddressBar, WM_SETTEXT, 0&, Text(0)
    
    modLogging.WriteLineToLog "IEAddressBar: Set text to: " & newText

End Property

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-06 08:05)  Decl: 4  Code: 152  Total: 156 Lines
':) CommentOnly: 13 (8.3%)  Commented: 0 (0%)  Filled: 98 (62.8%)  Empty: 58 (37.2%)  Max Logic Depth: 3
