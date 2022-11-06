VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents m_Browser As WebBrowser
Attribute m_Browser.VB_VarHelpID = -1

Private m_History As IUrlHistoryStg

Public Event NavigationIntercepted(destinationURL As String)

Public Sub Construct(hWndInternetExplorerServer As Long)

  ' VB6 classes can't have a real constructor. *sigh*

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
    IIDFromString StrPtr(IIDSTR_IWebBrowserApp), iid_SWebBrowserApp
    IIDFromString StrPtr(IIDSTR_IHTMLElement), iid_IHTMLElement
    IIDFromString StrPtr(IIDSTR_IWebBrowser2), iid_IWebBrowser2

    ' Documentation for what follows (sort of):
    ' http://jeffcode.blogspot.com/2008/06/internet-explorer-object-cookbook.html
    ' https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/249232
    ' https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/257717

    ' The approach below does it using AccessibleObjectFromWindow (Win98+ required)

    If AccessibleObjectFromWindow(hWndInternetExplorerServer, OBJID_CLIENT, iid_IDispatch, objIEBrowser) <> 0 Then
        modLogging.WriteLineToLog "IEBrowser: AccessibleObjectFromWindow failed."
        Exit Sub
    End If

    Set objIServiceProvider = objIEBrowser.accChild(CVar(0&))

    objIServiceProvider.QueryService iid_IHTMLElement, iid_IDispatch, objHTMLElement

    Set objIServiceProvider2 = objHTMLElement.Document

    objIServiceProvider2.QueryService iid_SWebBrowserApp, iid_IWebBrowser2, objWebBrowser

    Set m_Browser = objWebBrowser

    Set objIServiceProvider2 = Nothing
    Set objIServiceProvider = Nothing

    modLogging.WriteLineToLog "IEBrowser: Success!"

End Sub

Public Sub Destroy()

    Set m_Browser = Nothing

End Sub

Private Sub m_Browser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    Cancel = True
    modLogging.WriteLineToLog "IEBrowser: Catch and redirect navigating to: " & CStr(URL)

    RaiseEvent NavigationIntercepted(CStr(URL))

End Sub

Public Sub PushIntoHistory(ByVal URL As String, ByVal title As String)

    If m_History Is Nothing Then
        Set m_History = New UrlHistory
    End If

    modLogging.WriteLineToLog "IEBrowser: Add to history URL=" & URL & "; Title=" & title

    m_History.ADDURL URL, title, STATURLFLAG_ISTOPLEVEL

End Sub

Public Sub SetTitle(newTitle As String)

    If m_Browser Is Nothing Then
        Exit Sub
    End If

    m_Browser.Document.title = newTitle

End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-06 08:05)  Decl: 7  Code: 91  Total: 98 Lines
':) CommentOnly: 9 (9.2%)  Commented: 0 (0%)  Filled: 61 (62.2%)  Empty: 37 (37.8%)  Max Logic Depth: 2
