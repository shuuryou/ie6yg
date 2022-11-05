VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements SSubTUP.ISubclass

Private m_hWndIEFrame As Long
Private m_hWndButtonToolbar As Long
Private m_hWndReBarWindow32 As Long

Private m_ToolbarIdCommandBACK As Integer
Private m_ToolbarIdCommandFORWARD As Integer
Private m_ToolbarIdCommandSTOP As Integer
Private m_ToolbarIdCommandREFRESH As Integer
Private m_ToolbarIdCommandHOME As Integer
Private m_ToolbarIdCommandMEDIA As Integer

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
Private Const TOOLBAR_IBITMAP_MEDIA As Byte = 11
'Private Const TOOLBAR_IBITMAP_HISTORY As Byte = 12

Public Enum ToolbarCommand
    CommandUnknown = 0
    CommandBack = 1
    CommandForward = 2
    CommandStop = 3
    CommandRefresh = 4
    CommandHome = 5
    CommandMedia = 6
End Enum
#If False Then ':) Line inserted by Formatter
Private CommandUnknown, CommandBack, CommandForward, CommandStop, CommandRefresh, CommandHome, CommandMedia ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Public Event ToolbarButtonPressed(command As ToolbarCommand)
Public Event ToolbarMenuRequested(command As ToolbarCommand)

Public Sub Construct(hWndIEFrame As Long)

    m_hWndIEFrame = hWndIEFrame

    FindToolbar

    If m_hWndButtonToolbar = -1 Or m_hWndReBarWindow32 = -1 Then
        Exit Sub
    End If

    AttachMessage Me, m_hWndButtonToolbar, TB_ENABLEBUTTON
    AttachMessage Me, m_hWndButtonToolbar, TB_ENABLEBUTTON + TB_IE6YG_SHIFT
    
    AttachMessage Me, m_hWndReBarWindow32, WM_COMMAND
    AttachMessage Me, m_hWndReBarWindow32, WM_NOTIFY

End Sub

Public Sub Destroy()

    If m_hWndButtonToolbar = -1 Or m_hWndReBarWindow32 = -1 Then
        Exit Sub
    End If

    DetachMessage Me, m_hWndButtonToolbar, TB_ENABLEBUTTON
    DetachMessage Me, m_hWndButtonToolbar, TB_ENABLEBUTTON + TB_IE6YG_SHIFT
    
    DetachMessage Me, m_hWndReBarWindow32, WM_COMMAND
    DetachMessage Me, m_hWndReBarWindow32, WM_NOTIFY

End Sub

Private Function EnumToToolbarCommandId(command As ToolbarCommand) As Long

    If command = CommandBack Then
        EnumToToolbarCommandId = m_ToolbarIdCommandBACK
        Exit Function
    End If

    If command = CommandForward Then
        EnumToToolbarCommandId = m_ToolbarIdCommandFORWARD
        Exit Function
    End If

    If command = CommandStop Then
        EnumToToolbarCommandId = m_ToolbarIdCommandSTOP
        Exit Function
    End If

    If command = CommandRefresh Then
        EnumToToolbarCommandId = m_ToolbarIdCommandREFRESH
        Exit Function
    End If

    If command = CommandHome Then
        EnumToToolbarCommandId = m_ToolbarIdCommandHOME
        Exit Function
    End If

    If command = CommandMedia Then
        EnumToToolbarCommandId = m_ToolbarIdCommandMEDIA
        Exit Function
    End If

    EnumToToolbarCommandId = -1

End Function

Private Sub FindToolbar()

  Dim hWndTemp As Long
  Dim hwndToolbar As Long

  Dim sctTbButton As TBBUTTON
  Dim intButtonCount As Integer

  Dim i As Integer, j As Integer

    If m_hWndIEFrame = -1 Then
        modLogging.WriteLineToLog "IEToolbar: m_hWndIEFrame is bad."
        m_hWndButtonToolbar = -1
        Exit Sub
    End If

    hWndTemp = FindWindowEx(m_hWndIEFrame, 0&, "WorkerA", vbNullString)

    If hWndTemp = 0 Then
        modLogging.WriteLineToLog "IEToolbar: WorkerA not found."
        
        hWndTemp = FindWindowEx(m_hWndIEFrame, 0&, "WorkerW", vbNullString)
    
        If hWndTemp = 0 Then
            modLogging.WriteLineToLog "IEToolbar: WorkerW also not found."
            m_hWndButtonToolbar = -1
            Exit Sub
        Else
            modLogging.WriteLineToLog "IEToolbar: WorkerW found instead."
        End If
    End If

    hWndTemp = FindWindowEx(hWndTemp, 0&, "ReBarWindow32", vbNullString)

    If hWndTemp = 0 Then
        modLogging.WriteLineToLog "IEToolbar: ReBarWindow32 not found."
        m_hWndButtonToolbar = -1
        Exit Sub
    End If

    modLogging.WriteLineToLog "IEToolbar: Found ReBarWindow32: " & Hex$(hWndTemp)
    m_hWndReBarWindow32 = hWndTemp

    hwndToolbar = 0&

NextToolbar:
    m_ToolbarIdCommandBACK = -1
    m_ToolbarIdCommandFORWARD = -1
    m_ToolbarIdCommandSTOP = -1
    m_ToolbarIdCommandREFRESH = -1
    m_ToolbarIdCommandHOME = -1
    m_ToolbarIdCommandMEDIA = -1

    hwndToolbar = FindWindowEx(hWndTemp, hwndToolbar, "ToolbarWindow32", vbNullString)

    If hwndToolbar = 0 Then
        modLogging.WriteLineToLog "IEToolbar: ToolbarWindow32 not found or all of them were wrong."
        m_hWndButtonToolbar = -1
        m_hWndReBarWindow32 = -1
        Exit Sub
    End If

    intButtonCount = SendMessage(hwndToolbar, TB_BUTTONCOUNT, 0, 0)

    For i = 0 To intButtonCount - 1 Step 1
        SendMessage hwndToolbar, TB_GETBUTTON, i, sctTbButton

        ' Another shitty hack. It would have been nice to use TB_GETBUTTONTEXT,
        ' but this consistenly crashed the BHO and IE unless there was a MsgBox
        ' in this loop to interrupt it every iteration. I guess it's a bug in
        ' COMCTL32?

        ' Don't know. Don't care. All of this is insane anyway, so I gave up
        ' after a day of trying to debug the crash. BHO's have a very large
        ' iBitmap value, so the following is reliable enough to find the button
        ' toolbar.

        ' It doesn't help that most BHO development went down 15-20 years ago
        ' and the knowledge has since become somewhat lost on the Internet.
        If sctTbButton.idCommand <> 0 And sctTbButton.iBitmap = TOOLBAR_IBITMAP_REFRESH Then
            m_hWndButtonToolbar = hwndToolbar
            modLogging.WriteLineToLog "IEToolbar: Found it: " & Hex$(hwndToolbar)

            ' Now go through the buttons a second time to find the interesting
            ' ones by their bitmap index. The bitmap index seems to be stable.
            ' The command IDs change randomly when changing toolbar settings.

            For j = 0 To intButtonCount - 1 Step 1
                SendMessage hwndToolbar, TB_GETBUTTON, j, sctTbButton

                Select Case sctTbButton.iBitmap
                  Case TOOLBAR_IBITMAP_BACK
                    modLogging.WriteLineToLog "IEToolbar: Found BACK button. Command ID: " & sctTbButton.idCommand
                    m_ToolbarIdCommandBACK = sctTbButton.idCommand
                  Case TOOLBAR_IBITMAP_FORWARD
                    modLogging.WriteLineToLog "IEToolbar: Found FORWARD button. Command ID: " & sctTbButton.idCommand
                    m_ToolbarIdCommandFORWARD = sctTbButton.idCommand
                  Case TOOLBAR_IBITMAP_STOP
                    modLogging.WriteLineToLog "IEToolbar: Found STOP button. Command ID: " & sctTbButton.idCommand
                    m_ToolbarIdCommandSTOP = sctTbButton.idCommand
                  Case TOOLBAR_IBITMAP_REFRESH
                    modLogging.WriteLineToLog "IEToolbar: Found REFRESH button. Command ID: " & sctTbButton.idCommand
                    m_ToolbarIdCommandREFRESH = sctTbButton.idCommand
                  Case TOOLBAR_IBITMAP_HOME:
                    modLogging.WriteLineToLog "IEToolbar: Found HOME button. Command ID: " & sctTbButton.idCommand
                    m_ToolbarIdCommandHOME = sctTbButton.idCommand
                  Case TOOLBAR_IBITMAP_MEDIA:
                    modLogging.WriteLineToLog "IEToolbar: Found MEDIA button. Command ID: " & sctTbButton.idCommand
                    m_ToolbarIdCommandMEDIA = sctTbButton.idCommand

                End Select
            Next j

            Exit Sub
        End If
    Next i

    GoTo NextToolbar

End Sub

Public Property Get hWndButtonToolbar() As Long

    hWndButtonToolbar = m_hWndButtonToolbar

End Property

Public Property Get hWndReBarWindow32() As Long

    hWndReBarWindow32 = m_hWndReBarWindow32

End Property

Private Property Get ISubclass_MsgResponse() As SSubTUP.EMsgResponse

    ISubclass_MsgResponse = emrConsume

End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTUP.EMsgResponse)

    ' Unused

End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim command As ToolbarCommand
  
    modLogging.WriteLineToLog "imsg:" & iMsg
  
    If hwnd = m_hWndButtonToolbar Then
        If iMsg = TB_ENABLEBUTTON Then
            Exit Function
        End If
        
        If iMsg = TB_ENABLEBUTTON + TB_IE6YG_SHIFT Then
            iMsg = iMsg - TB_IE6YG_SHIFT
        End If
        
        ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
        Exit Function
    End If

    If hwnd = m_hWndReBarWindow32 Then
        If iMsg = WM_COMMAND Then
            command = ToolbarCommandIdToEnum(wParam)

            If command <> CommandUnknown Then
                RaiseEvent ToolbarButtonPressed(command)
            Else
                ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            End If

            Exit Function
        End If
        
        If iMsg = WM_NOTIFY Then
            ' User clicked the little dropdown button on the back or forward
            ' button. Luckily forwarding this as an event, so the frontend
            ' can call the VB6 PopupMenu function, works perfectly. No extra
            ' effort needed (other than having to implement synced history
            ' menus *sigh*).
            
            Dim nmTB As NMTOOLBAR_SHORT
            CopyMemory nmTB, ByVal lParam, Len(nmTB)
            
            If nmTB.hdr.code = TBN_DROPDOWN Then
                modLogging.WriteLineToLog CStr(nmTB.iItem)
                
                If nmTB.iItem = m_ToolbarIdCommandBACK Then
                    RaiseEvent ToolbarMenuRequested(ToolbarCommand.CommandBack)
                    Exit Function
                End If
                
                If nmTB.iItem = m_ToolbarIdCommandFORWARD Then
                    RaiseEvent ToolbarMenuRequested(ToolbarCommand.CommandForward)
                    Exit Function
                End If
            End If
        End If

        ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
        Exit Function
    End If

End Function

Public Sub SetToolbarCommandState(command As ToolbarCommand, enable As Boolean)

  Dim commandId As Long

    commandId = EnumToToolbarCommandId(command)

    If commandId = -1 Then
        modLogging.WriteLineToLog "IEToolbar: Cannot set state on unknown command."
        Exit Sub
    End If

    SendMessage m_hWndButtonToolbar, TB_ENABLEBUTTON + TB_IE6YG_SHIFT, ByVal commandId, ByVal loWord(IIf(enable, 1, 0))
    modLogging.WriteLineToLog "SetToolbarCommandState: Command ID " & commandId & " new state: " & enable

End Sub

Private Function ToolbarCommandIdToEnum(idCommand As Long) As ToolbarCommand

    If idCommand = m_ToolbarIdCommandBACK Then
        ToolbarCommandIdToEnum = CommandBack
        Exit Function
    End If

    If idCommand = m_ToolbarIdCommandFORWARD Then
        ToolbarCommandIdToEnum = CommandForward
        Exit Function
    End If

    If idCommand = m_ToolbarIdCommandSTOP Then
        ToolbarCommandIdToEnum = CommandStop
        Exit Function
    End If

    If idCommand = m_ToolbarIdCommandREFRESH Then
        ToolbarCommandIdToEnum = CommandRefresh
        Exit Function
    End If

    If idCommand = m_ToolbarIdCommandHOME Then
        ToolbarCommandIdToEnum = CommandHome
        Exit Function
    End If

    If idCommand = m_ToolbarIdCommandMEDIA Then
        ToolbarCommandIdToEnum = CommandMedia
        Exit Function
    End If

    ToolbarCommandIdToEnum = CommandUnknown

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-03 00:39)  Decl: 51  Code: 298  Total: 349 Lines
':) CommentOnly: 24 (6,9%)  Commented: 3 (0,9%)  Filled: 259 (74,2%)  Empty: 90 (25,8%)  Max Logic Depth: 5
