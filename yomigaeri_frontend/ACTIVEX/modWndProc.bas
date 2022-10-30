Attribute VB_Name = "modWndProc"
Option Explicit

' This module is required because WndProc implementations must live in a 
' global module and cannot be a member of an instantiated class. This is 
' also why the IEBrowserManager class has so many weird properties. VB6 
' is not a nice language to program in. Alas. 

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_USER As Long = &H400

Private Const WM_COMMAND As Long = &H111

Private Const TB_ENABLEBUTTON As Long = (WM_USER + &H1)

Private Const SB_SETTEXT As Long = WM_USER + 1
Private Const SB_SETPARTS As Long = WM_USER + 4
Private Const SB_SETMINHEIGHT As Long = WM_USER + 8
Private Const SB_SETICON As Long = WM_USER + 15


Public BROWSER_MANAGER_INSTANCE As IEBrowserManager

Public Function ButtonToolbarWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If uMsg = TB_ENABLEBUTTON Then
        Select Case wParam
          Case BROWSER_MANAGER_INSTANCE.ToolbarIdCommandBack:
            lParam = Abs(BROWSER_MANAGER_INSTANCE.ToolbarButtonStateBack)
          Case BROWSER_MANAGER_INSTANCE.ToolbarIdCommandForward:
            lParam = Abs(BROWSER_MANAGER_INSTANCE.ToolbarButtonStateForward)
          Case BROWSER_MANAGER_INSTANCE.ToolbarIdCommandStop:
            lParam = Abs(BROWSER_MANAGER_INSTANCE.ToolbarButtonStateStop)
          Case BROWSER_MANAGER_INSTANCE.ToolbarIdCommandRefresh:
            lParam = Abs(BROWSER_MANAGER_INSTANCE.ToolbarButtonStateRefresh)
        End Select
    End If

    ButtonToolbarWndProc = CallWindowProc(BROWSER_MANAGER_INSTANCE.ToolbarOldWndProc, hWnd, uMsg, wParam, lParam)

End Function

Public Function RebarWindow32WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If uMsg = WM_COMMAND Then
        Select Case wParam
          Case BROWSER_MANAGER_INSTANCE.ToolbarIdCommandBack:
          Case BROWSER_MANAGER_INSTANCE.ToolbarIdCommandForward:
          Case BROWSER_MANAGER_INSTANCE.ToolbarIdCommandStop:
          Case BROWSER_MANAGER_INSTANCE.ToolbarIdCommandRefresh:
            BROWSER_MANAGER_INSTANCE.RaiseButtonToolbarEvent wParam
            Exit Function '---> Bottom
        End Select
    End If

    RebarWindow32WndProc = CallWindowProc(BROWSER_MANAGER_INSTANCE.ReBarOldWndProc, hWnd, uMsg, wParam, lParam)

End Function

Public Function StatusBarWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg
        Case SB_SETTEXT:
        Case SB_SETPARTS:
        lParam = 4
        Case SB_SETMINHEIGHT:
        Case SB_SETICON:
                Exit Function
    End Select

    StatusBarWndProc = CallWindowProc(BROWSER_MANAGER_INSTANCE.StatusBarOldWndProc, hWnd, uMsg, wParam, lParam)

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Oct-29 22:40)  Decl: 9  Code: 44  Total: 53 Lines
':) CommentOnly: 6 (11.3%)  Commented: 0 (0%)  Filled: 40 (75.5%)  Empty: 13 (24.5%)  Max Logic Depth: 3
