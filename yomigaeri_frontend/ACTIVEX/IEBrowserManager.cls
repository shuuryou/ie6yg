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

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_SETTEXT As Long = &HC

Private Declare Function AccessibleObjectFromWindow Lib "oleacc.dll" (ByVal hWnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long
Private Const OBJID_CLIENT As Long = -4

Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long
Private Const IIDSTR_IHTMLElement As String = "{3050f1ff-98b5-11cf-bb82-00aa00bdce0b}"
Private Const IIDSTR_IWebBrowser2 As String = "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"
Private Const IIDSTR_SWebBrowserApp As String = "{0002DF05-0000-0000-C000-000000000046}"

Private m_History As IUrlHistoryStg
Private WithEvents m_Browser As WebBrowser
Attribute m_Browser.VB_VarHelpID = -1

Private m_hWndUserControl As Long
Private m_hWndAddressBar As Long

Public Event IEWantsToNavigate(newUrl As String)

Public Property Let hWndUserControl(ByRef hWnd As Long)
    m_hWndUserControl = hWnd
End Property

Public Property Get hWndUserControl() As Long
    hWndUserControl = m_hWndUserControl
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

    modLogging.WriteLineToLog "Internet Explorer_Server hWnd is " & Hex$(hWnd)

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

    modLogging.WriteLineToLog "Shell DocObject View hWnd is " & Hex$(hWnd)

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

    modLogging.WriteLineToLog "IEFrame hWnd is " & Hex$(hWnd)

    FindIEFrame = hWnd

End Function

Private Function FindIE6AddressBarHandle() As Long

  Dim hWnd As Long

    hWnd = FindIEFrame()

    If hWnd = -1 Then
        Exit Function '---> Bottom
    End If

    ' ------------------------------------------------------------------------

    ' Now start going down again and find the first "WorkerW"

    hWnd = FindWindowEx(hWnd, 0&, "WorkerW", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindWindowEx WorkerW failed. GetLastError returns " & Hex$(Err.LastDllError)
        FindIE6AddressBarHandle = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "WorkerW hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the "ReBarWindow32" containing the address bar

    hWnd = FindWindowEx(hWnd, 0&, "ReBarWindow32", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindWindowEx ReBarWindow32 failed. GetLastError returns " & Hex$(Err.LastDllError)
        FindIE6AddressBarHandle = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "ReBarWindow32 hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the "ComboBoxEx32" to get to the address bar

    hWnd = FindWindowEx(hWnd, 0&, "ComboBoxEx32", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindWindowEx ComboBoxEx32 failed. GetLastError returns " & Hex$(Err.LastDllError)
        FindIE6AddressBarHandle = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "ComboBoxEx32 hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the actual "ComboBox" to get to the edit control

    hWnd = FindWindowEx(hWnd, 0&, "ComboBox", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindWindowEx ComboBox failed. GetLastError returns " & Hex$(Err.LastDllError)
        FindIE6AddressBarHandle = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "ComboBox hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Finally find the "Edit" window and we're there

    hWnd = FindWindowEx(hWnd, 0&, "Edit", vbNullString)

    If Err.LastDllError <> 0 Or hWnd = 0 Then
        modLogging.WriteLineToLog "FindWindowEx Edit failed. GetLastError returns " & Hex$(Err.LastDllError)
        FindIE6AddressBarHandle = -1
        Exit Function '---> Bottom
    End If

    modLogging.WriteLineToLog "Edit hWnd is " & Hex$(hWnd)

    FindIE6AddressBarHandle = hWnd

End Function

Public Sub PushIntoHistory(ByVal URL As String, ByVal title As String)
    If m_History Is Nothing Then
        Set m_History = New UrlHistory
    End If
    
    m_History.ADDURL URL, title, STATURLFLAG_ISTOPLEVEL
End Sub

Public Function GetAddressBarHandle() As Long

  Dim ret As Long

    modLogging.WriteLineToLog "Try to find IE6 address bar handle from " & Hex$(hWndUserControl)

    ret = FindIE6AddressBarHandle()

    If ret <> -1 Then
        GetAddressBarHandle = ret
        Exit Function '---> Bottom
    End If

    ' If we implement support for other browsers like IE5 or IE7 in the
    ' future, try them here with the above pattern repeated.

    GetAddressBarHandle = -1

End Function

Public Sub SetAddressBarText(newText As String)

    If m_hWndAddressBar = 0 Then
        m_hWndAddressBar = GetAddressBarHandle()
    End If
    
    If m_hWndAddressBar = -1 Then
        modLogging.WriteLineToLog "SetAddressBarText: Cannot set if address bar not found."
        Exit Sub
    End If

    modLogging.WriteLineToLog "SetAddressBarText: Text becomes: " & newText

    SendMessage m_hWndAddressBar, WM_SETTEXT, 0&, ByVal newText

End Sub

Public Sub LockOntoIWebBrowser2()
    
    Dim iid_IDispatch As UUID
    Dim iid_SWebBrowserApp As UUID
    Dim iid_IHTMLElement As UUID
    Dim iid_IWebBrowser2 As UUID
    
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
    
    Dim hWndInternetExplorerServer As Long
    hWndInternetExplorerServer = FindInternetExplorerServer()
    
    If hWndInternetExplorerServer = -1 Then
        modLogging.WriteLineToLog "BlockNavigation: hWnd not found."
        Exit Sub
    End If
    
    ' Documentation for what follows (sort of):
    ' http://jeffcode.blogspot.com/2008/06/internet-explorer-object-cookbook.html
    ' https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/249232
    ' https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/257717
    
    ' The approach below does it using AccessibleObjectFromWindow (Win98+ required)

    If AccessibleObjectFromWindow(hWndInternetExplorerServer, OBJID_CLIENT, iid_IDispatch, objIEBrowser) <> 0 Then
        modLogging.WriteLineToLog "BlockNavigation: AccessibleObjectFromWindow failed."
        Exit Sub
    End If
 
    Set objIServiceProvider = objIEBrowser.accChild(CVar(0&))
 
    objIServiceProvider.QueryService iid_IHTMLElement, iid_IDispatch, objHTMLElement
    
    Set objIServiceProvider2 = objHTMLElement.Document
   
    objIServiceProvider2.QueryService iid_SWebBrowserApp, iid_IWebBrowser2, objWebBrowser
    
    Set m_Browser = objWebBrowser
    
    Set objIServiceProvider2 = Nothing
    Set objIServiceProvider = Nothing
    
    modLogging.WriteLineToLog "BlockNavigation: Succeeded getting IWebBrowser2!"
End Sub

Public Sub SetStatusBarText(text As String)
    ' This is unfortunately not reliable very early during ActiveX control
    ' initialization when the page has not yet settled. Therefore it is
    ' done using a JScript event instead.

    If m_Browser Is Nothing Then
        modLogging.WriteLineToLog "SetIEStatusBar: IWebBrowser2 not locked onto."
        Exit Sub
    End If
    
    m_Browser.StatusText = text
    modLogging.WriteLineToLog "SetIEStatusBar: Text becomes: " & text
End Sub

Private Sub m_Browser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Cancel = True
    modLogging.WriteLineToLog "Prevent IE from navigating to this URL: " & CStr(URL)
    RaiseEvent IEWantsToNavigate(CStr(URL))
End Sub

Private Sub m_Browser_DownloadBegin()
    m_Browser.Stop
End Sub
