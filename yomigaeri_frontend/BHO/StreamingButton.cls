VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamingButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Adapted from the book
' Visual Basic Shell Programming
' J.P.Hamilton
' Publisher: O 'Reilly
' First Edition July 2000
' ISBN: 1-56592-670-6

Implements IObjectWithSite
Implements IOleCommandTarget

Private m_pUnkSite As IUnknownVB
Private WithEvents m_IEBrowser As InternetExplorer
Attribute m_IEBrowser.VB_VarHelpID = -1

Private Const IID_IWebBrowserApp = "{0002DF05-0000-0000-C000-000000000046}"
Private Const IID_IWebBrowser2 = "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"

Private Sub IObjectWithSite_GetSite(ByVal priid As VBShellLib.REFIID, ppvObj As VBShellLib.VOID)

    m_pUnkSite.QueryInterface priid, ppvObj

End Sub

Private Sub IObjectWithSite_SetSite(ByVal pSite As VBShellLib.IUnknownVB)

  Dim pServiceProvider As IServiceProvider
  Dim clsidWebApp As GUID
  Dim clsidWebBrowser2 As GUID

    If ObjPtr(pSite) = 0 Then
        CopyMemory m_IEBrowser, 0&, 4
        Exit Sub '--->�Bottom
    End If

    'Save the site pointer for GetSite
    Set m_pUnkSite = pSite

    Set pServiceProvider = m_pUnkSite

    'Query service provider to get IWebBrowser2 (InternetExplorer)
    CLSIDFromString StrPtr(IID_IWebBrowserApp), clsidWebApp
    CLSIDFromString StrPtr(IID_IWebBrowser2), clsidWebBrowser2

    Set m_IEBrowser = pServiceProvider.QueryService(VarPtr(clsidWebApp), VarPtr(clsidWebBrowser2))

    Set pServiceProvider = Nothing

End Sub

Private Sub IOleCommandTarget_Exec(ByVal pguidCmdGroup As VBShellLib.LPGUID, ByVal nCmdID As VBShellLib.DWORD, ByVal nCmdExecOpt As VBShellLib.DWORD, pvaIn As Variant, pvaOut As Variant)

  ' In order to invoke a Component Object Model (COM) object from Internet Explorer,
  ' it must implement IOleCommandTarget. Only one command is supported per object;
  ' the COM object's IOleCommandTarget::Exec is always called with nCmdID=0 and with
  ' VARIANT arguments set to NULL. Additionally, the implementation of
  ' IOleCommandTarget::QueryStatus is always called with cCmds=1.

  ' https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa753588(v=vs.85)

  Dim pDoc As IHTMLDocument2
  Dim objElement As IHTMLElement

    Set pDoc = m_IEBrowser.document

    For Each objElement In pDoc.All
        If objElement.getAttribute("classid", 0) = "clsid:D322D3BD-AF48-4787-ACA6-2D32F2A59A32" Then
            objElement.QueryStreamingAvailable
            Exit Sub '--->�Bottom
        End If
    Next objElement
    
    ' 101: Web streaming support is not available for the current website.
    ' 102: Web Streaming Support
    MsgBox LoadResString(101), vbCritical, LoadResString(102)
End Sub

Private Sub IOleCommandTarget_QueryStatus(ByVal pguidCmdGroup As VBShellLib.LPGUID, ByVal cCmds As VBShellLib.ULONG, ByVal prgCmds As VBShellLib.LPOLECMD, ByVal pCmdText As VBShellLib.LPOLECMDTEXT)

  Dim i As Integer
  Dim cmd As OLECMD

    For i = 0 To cCmds - 1
        CopyMemory cmd, ByVal prgCmds + (Len(cmd) * i), Len(cmd)
        cmd.cmdf = OLECMDF_SUPPORTED Or OLECMDF_ENABLED
        CopyMemory ByVal prgCmds + (Len(cmd) * i), cmd, Len(cmd)
    Next i

End Sub

':) Ulli's VB Code Formatter V2.24.17 (2022-Oct-26 21:22)  Decl: 17  Code: 72  Total: 89 Lines
':) CommentOnly: 16 (18%)  Commented: 0 (0%)  Filled: 61 (68.5%)  Empty: 28 (31.5%)  Max Logic Depth: 3