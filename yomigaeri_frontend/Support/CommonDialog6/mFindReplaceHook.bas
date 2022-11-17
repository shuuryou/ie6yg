Attribute VB_Name = "mFindReplaceHook"
Option Explicit

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Const WH_KEYBOARD = 2
Private Const WH_MSGFILTER = (-1)

Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type Msg
   hWnd As Long
   message As Long
   wParam As Long
   lParam As Long
   time As Long
   pt As POINTAPI
End Type
Private Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As Msg) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetFocusAPI Lib "user32" Alias "GetFocus" () As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetNextDlgGroupItem Lib "user32" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
Private Declare Function GetNextDlgTabItem Lib "user32" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
Private Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101

Private m_hHook As Long
Private m_hWnd As Long

Public Sub Attach(ByVal hWnd As Long)
   If (m_hHook = 0) Then
      Dim lpFn As Long
      lpFn = HookProcAddress(AddressOf HookProc)
      m_hHook = SetWindowsHookEx(WH_KEYBOARD, lpFn, App.hInstance, GetCurrentThreadId())
      If Not m_hHook = 0 Then
         m_hWnd = hWnd
      Else
         Debug.Assert (m_hHook <> 0)
      End If
   End If
End Sub
Public Sub Detach()
   If Not (m_hHook = 0) Then
      UnhookWindowsHookEx m_hHook
      m_hHook = 0
      m_hWnd = 0
   End If
End Sub

Private Function HookProcAddress(ByVal lPtr As Long) As Long
   HookProcAddress = lPtr
End Function


Private Function HookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lR As Long
Dim bKeyDown As Boolean
   If (nCode >= 0) Then
      If (GetActiveWindow() = m_hWnd) Then
         Dim lpMsg As Msg
         lpMsg.hWnd = m_hWnd
         lpMsg.time = GetCurrentTime
         Dim pt As POINTAPI
         GetCursorPos pt
         lpMsg.pt = pt
         lpMsg.lParam = lParam
         lpMsg.wParam = wParam
         lpMsg.message = IIf((lParam And &H40000000) = &H40000000, WM_KEYUP, WM_KEYDOWN)
         If (wParam = vbKeyEscape) Or (wParam = vbKeyReturn) Or (wParam = vbKeyTab) Then
            If Not (IsDialogMessage(m_hWnd, lpMsg) = 0) Then
               lR = 1
            End If
         End If
      End If
      If (lR = 1) Then
         HookProc = 1
      Else
         HookProc = CallNextHookEx(m_hHook, nCode, wParam, lParam)
      End If
   End If
   Exit Function
   
End Function
