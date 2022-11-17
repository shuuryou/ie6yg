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

Public Sub AddToHistory(ByVal url As String, ByVal title As String)

    If m_History Is Nothing Then
        Set m_History = New UrlHistory
    End If

    If url = "" Then
        modLogging.WriteLineToLog "IEBrowser: Refuse to add history with empty URL."
        Exit Sub
    End If

    ' IUrlHistoryStg.AddUrl does not work with HTTPS URLs, but without access
    ' to the Windows source code you'll *never* learn the real reason for it.
    ' It's not documented anywhere; not even in ancient MSDN Library releases.

    ' The reason can be found in leaked WinXP code (and IE5 Bug #110378)

    ' File: shell/shdocvw/urlhist.cpp
    ' Function: HRESULT CUrlHistory::AddUrlAndNotifyCP

    ' "For secure URLs (https:) [...] we presume that these can contain
    '  sensitive data that the site did not want retained on the client
    '  machine. Therefore, we do not write a history entry for such sites."

    ' Horseshit. Well, back in 1999 that was probably valid thinking. But in
    ' 2022, everyone uses HTTPS for everything thanks to NSA and Snowden.

    ' Workaround: Fool IE6 into thinking the HTTPS website is already cached.

    If InStr(1, UCase$(url), "HTTPS://", vbBinaryCompare) = 1 Then
  Dim lngBufSize As Long

  Dim strHost As String
        strHost = Space$(INTERNET_MAX_URL_LENGTH)
        lngBufSize = Len(strHost)

        UrlGetPartA url, strHost, lngBufSize, URL_PART_HOSTNAME, 0&

        If lngBufSize = INTERNET_MAX_PATH_LENGTH Or lngBufSize = 0 Then
            modLogging.WriteLineToLog "IEBrowse: Can't add history because UrlGetPartA failed (1)."
            Exit Sub
        End If

        strHost = Mid$(strHost, 1, lngBufSize)

  Dim strTempURL As String
        strTempURL = "https://" & strHost & "/ie6yg.htm"

  Dim strFileName As String
        strFileName = Space$(MAX_PATH)

        CreateUrlCacheEntry strTempURL, 0, "html", strFileName, 0&

        If Err.LastDllError <> 0 Then
            modLogging.WriteLineToLog "IEBrowse: Can't add history because CreateUrlCacheEntry failed. HRESULT: " & Hex$(Err.LastDllError)
            Exit Sub
        End If

        ' MSDN Library January 2001:
        ' "Clients that add entries to the cache should set the headers to at
        '  least "HTTP/1.0 200 OK\r\n\r\n"; otherwise, Internet Explorer and
        '  other client applications that are utilizing the Win32 Internet
        '  functions will disregard the entry."

        ' That's perfect, since the entries made by IE6YG are fake garbage.
        ' Since we do not pass a pointer to INTERNET_CACHE_ENTRY_INFO below,
        ' no headers are set. But it's enough to fool IE6 into thinking that
        ' the HTTPS website passed in via "url" didn't set cache headers and
        ' allows adding the URL to the history to succeed.

  Dim ftExpireTime As FILETIME
  Dim ftModifiedTime As FILETIME

  Dim lngRet As Long

        lngRet = CommitUrlCacheEntry(url, strFileName, ftExpireTime, _
                 ftModifiedTime, NORMAL_CACHE_ENTRY, 0&, 0&, 0&, 0&)

        If lngRet <> 1 Then ' False
            ' Sometimes HRESULT is 0x20, but it still works?!?!?
            modLogging.WriteLineToLog "IEBrowse: CommitUrlCacheEntry failed, but continue anway. HRESULT: " & Hex$(Err.LastDllError)
            'Exit Sub
        End If
    End If

    m_History.ADDURL url, title, 0&
    modLogging.WriteLineToLog "IEBrowse: Add to history URL=" & url & "; Title=" & title

End Sub

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

Private Sub m_Browser_BeforeNavigate2(ByVal pDisp As Object, url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    Cancel = True
    modLogging.WriteLineToLog "IEBrowser: Catch and redirect navigating to: " & CStr(url)

    RaiseEvent NavigationIntercepted(CStr(url))

End Sub

Public Sub SetTitle(newTitle As String)

    If m_Browser Is Nothing Then
        Exit Sub
    End If

    m_Browser.Document.title = newTitle

End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-17 19:43)  Decl: 7  Code: 169  Total: 176 Lines
':) CommentOnly: 33 (18.8%)  Commented: 1 (0.6%)  Filled: 118 (67%)  Empty: 58 (33%)  Max Logic Depth: 3
