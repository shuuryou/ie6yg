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

Private Declare Function EnableWindow Lib "USER32.DLL" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Private Type COMBOBOXEXITEM
    mask As Long ' CBEIF..
    iItem As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    iOverlay As Long
    iIndent As Long
    lParam As Long
End Type

Private Const CB_ERR = (-1)
Private Const CB_GETCOUNT = &H146
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158

Private Const CBEIF_TEXT = &H1
Private Const CBEIF_IMAGE = &H2
Private Const CBEIF_SELECTEDIMAGE = &H4

Private Const CBEM_SETITEMA = (WM_USER + 5)
Private Const CBEM_GETITEMA = (WM_USER + 4)
Private Const CBEM_DELETEITEM = &H144
Private Const CBEM_INSERTITEMA = (WM_USER + 1)

Private m_hWndIEFrame As Long
Private m_hWndComboBoxEx As Long

Private m_CBEFImage As Long
Private m_CBEFSelectedImage As Long

Private m_Enabled As Boolean

Public Property Let Address(newAddress As String)

  Dim strAddress As String

    If m_hWndComboBoxEx = -1 Then
        Exit Property
    End If

    strAddress = StrConv(newAddress, vbFromUnicode)

    ' From leaked WinXP SP1 source code.
    ' File: shell/browseui/addrlist.cpp
    ' Function: HRESULT CAddressList::NavigationComplete

    ' Chromium takes care of making URLs look nice, so all the URL cleanup
    ' code (SHRemoveURLTurd, SHCleanupUrlForDisplay, ...) isn't used here.

  Dim item As COMBOBOXEXITEM

    With item
        .mask = CBEIF_TEXT Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE
        .pszText = strAddress
        .cchTextMax = LenB(strAddress)
        .iImage = m_CBEFImage
        .iSelectedImage = m_CBEFSelectedImage
        .iItem = -1
    End With

    If SendMessage(m_hWndComboBoxEx, CBEM_SETITEMA, 0, item) = 0 Then
        modLogging.WriteLineToLog "IEAddressBar: Failed to set text to: " & strAddress
        Exit Property
    End If

    modLogging.WriteLineToLog "IEAddressBar: Successfully set text to: " & strAddress

    ' Do not pass "strAddress" since it was already converted to
    ' an ANSI string by StrConv above. ComboBoxInsertURL also
    ' wants to do that, and if you StrConv a StrConv'ed string,
    ' it turns into question marks.

    ComboBoxInsertURL newAddress, 15

End Property

Public Property Get Address() As String

  Dim strAddress As String

    If m_hWndComboBoxEx = -1 Then
        Exit Property
    End If

    strAddress = Space$(INTERNET_MAX_URL_LENGTH + 1)

  Dim item As COMBOBOXEXITEM

    With item
        .mask = CBEIF_TEXT
        .pszText = strAddress
        .cchTextMax = LenB(strAddress)
        .iItem = -1
    End With

    If SendMessage(m_hWndComboBoxEx, CBEM_GETITEMA, 0, item) = 0 Then
        Address = ""
        Exit Property
    End If

    Address = TrimNull(StrConv(item.pszText, vbUnicode))

End Property

Private Sub ComboBoxInsertURL(szURL As String, nMaxComboBoxSize As Long)

  ' Translated from leaked WinXP SP1 source code.
  ' File: shell/browseui/addrlist.cpp

  ' /*******************************************************************
  '     FUNCTION: _ComboBoxInsertURL
  '     DESCRIPTION:
  '         Adds the specified URL to the top of the address bar
  '     combo box.  Limits the number of URLs in combo box to
  '     nMaxComboBoxSize.
  ' ********************************************************************/

  Dim strURLConvForSendMessage() As Byte
  Dim strURLConv As String

    strURLConvForSendMessage = StrConv(szURL, vbFromUnicode) & Chr$(0)
    strURLConv = StrConv(szURL, vbFromUnicode)

    ' // Since we own it and it's populated,
    ' // we will add it directly to the ComboBox.
  Dim iPrevInstance As Long

  Dim cbexItem As COMBOBOXEXITEM

    cbexItem.mask = CBEIF_TEXT Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE
    cbexItem.iItem = 0

    cbexItem.cchTextMax = LenB(strURLConv)
    cbexItem.pszText = strURLConv

    cbexItem.iImage = m_CBEFImage
    cbexItem.iSelectedImage = m_CBEFSelectedImage

    iPrevInstance = SendMessage(m_hWndComboBoxEx, CB_FINDSTRINGEXACT, -1&, strURLConvForSendMessage(0))

    ' Microsoft did not check if iPrevInstance is not 0. That's an optimization from IE6YG.
    If iPrevInstance <> CB_ERR And iPrevInstance <> 0 Then
        MoveAddressToTopOfList iPrevInstance
        Exit Sub
    End If

    ' // insert the URL as the first item in combo box
    SendMessage m_hWndComboBoxEx, CBEM_INSERTITEMA, 0&, cbexItem

    If SendMessage(m_hWndComboBoxEx, CB_GETCOUNT, 0&, 0&) > nMaxComboBoxSize Then
        ' // if we're ever over the limit, we should only be over the limit
        ' // by exactly one item
        If SendMessage(m_hWndComboBoxEx, CB_GETCOUNT, 0&, 0&) <> nMaxComboBoxSize + 1 Then
            modLogging.WriteLineToLog "IEAddressBar: ComboBoxInsertURL: ASSERTION FAILED!"
        End If

        ' // if over the limit, delete the least recently used
        ' // (the one with the highest index)

        SendMessage m_hWndComboBoxEx, CBEM_DELETEITEM, nMaxComboBoxSize, 0&
    End If

End Sub

Public Sub Construct(hWndIEFrame As Long)

  ' VB6 classes can't have a real constructor. *sigh*

    m_hWndIEFrame = hWndIEFrame

    FindAddressBar

    FindCBEXiImages

    ' The ComboBox is only populated with the item shown in the Edit control
    ' at the beginning (CMRUList::_PopulateOneItem). CMRUList::_Populate is
    ' called as soon as the user attempts to open the list of other items by
    ' a WM_COMMAND handler for CBN_DROPDOWN (CAddressList::_OnCommand).
    ' This is bad. All of the address bar shenanigans done by this class get
    ' a little bit messed up when IE places its own items into the ComboBox.

    ' The workaround is to coax IE into populating the dropdown immediately.
    ' We send a fake CBN_DROPDOWN with bogus handles. That's fine though,
    ' since the handling code inside IE doesn't care about the handles.

    SendMessage m_hWndComboBoxEx, WM_COMMAND, MAKELONG(0&, 7), 0& ' 7 = CBN_DROPDOWN

    ' If the following isn't sent, the backspace key no longer works.

    SendMessage m_hWndComboBoxEx, WM_COMMAND, MAKELONG(0&, 8), 0& ' 7 = CBN_CLOSEUP

End Sub

Public Property Let Enabled(enable As Boolean)

    If m_hWndComboBoxEx = -1 Then
        Exit Property
    End If

    m_Enabled = enable

    EnableWindow m_hWndComboBoxEx, enable

End Property

Public Property Get Enabled() As Boolean

    Enabled = m_Enabled

End Property

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
            m_hWndComboBoxEx = -1
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
        m_hWndComboBoxEx = -1
        Exit Sub
    End If

    modLogging.WriteLineToLog "IEAddressBar: ReBarWindow32 hWnd is " & Hex$(hWnd)

    ' ------------------------------------------------------------------------

    ' Now find the "ComboBoxEx32" to get to the address bar

    hWnd = FindWindowEx(hWnd, 0&, "ComboBoxEx32", vbNullString)

    If hWnd = 0 Then
        modLogging.WriteLineToLog "IEAddressBar: FindWindowEx ComboBoxEx32 failed."
        m_hWndComboBoxEx = -1
        Exit Sub
    End If

    modLogging.WriteLineToLog "IEAddressBar: ComboBoxEx32 hWnd is " & Hex$(hWnd)

    m_hWndComboBoxEx = hWnd

    modLogging.WriteLineToLog "IEAddressBar: Success!"

End Sub

Private Sub FindCBEXiImages()

  Dim item As COMBOBOXEXITEM

    If m_hWndComboBoxEx = -1 Then
        Exit Sub
    End If

    item.mask = CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE
    item.iItem = -1

    If SendMessage(m_hWndComboBoxEx, CBEM_GETITEMA, 0, item) = 0 Then
        modLogging.WriteLineToLog "IEAddressBar: Can't find image indices. CBEM_GETITEMA failed."
        m_CBEFImage = -1
        m_CBEFSelectedImage = -1
    End If

    m_CBEFImage = item.iImage
    m_CBEFSelectedImage = item.iSelectedImage

    modLogging.WriteLineToLog "IEAddressBar: Image indices are: iImage: " & m_CBEFImage & ", iSelectedImage: " & m_CBEFSelectedImage

End Sub

Public Property Get hWnd() As Long

    hWnd = m_hWndComboBoxEx

End Property

Private Function MoveAddressToTopOfList(iSel As Long) As Boolean

  ' Translated from leaked WinXP SP1 source code.
  ' File: shell/browseui/addrlist.cpp

  ' /*******************************************************************
  ' FUNCTION: _MoveAddressToTopOfList
  ' PARAMETERS:
  '         iSel - index of item in combo box to move
  ' DESCRIPTION:
  '         Moves the specified selection in the combo box
  ' to be the first item in the combo box
  ' ********************************************************************/

  Dim fRet As Boolean

    If iSel < 0 Then
        Exit Function ' // must have valid index
    End If

  Dim cbexItem As COMBOBOXEXITEM
  Dim szAddress As String
    szAddress = Space$(INTERNET_MAX_URL_LENGTH + 1)

    cbexItem.mask = CBEIF_TEXT Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE
    cbexItem.pszText = szAddress
    cbexItem.cchTextMax = LenB(szAddress)
    cbexItem.iItem = iSel

    If SendMessage(m_hWndComboBoxEx, CBEM_GETITEMA, 0, cbexItem) <> 0 Then

        SendMessage m_hWndComboBoxEx, CBEM_DELETEITEM, iSel, 0&

        cbexItem.iItem = 0

        ' // sending CBEM_INSERTITEM should return the index we specified
        ' // (0) if successful
        fRet = (SendMessage(m_hWndComboBoxEx, CBEM_INSERTITEMA, 0, cbexItem) = 0)
    End If

    MoveAddressToTopOfList = fRet

End Function

':) Ulli's VB Code Formatter V2.24.17 (2022-Dec-07 09:09)  Decl: 37  Code: 321  Total: 358 Lines
':) CommentOnly: 57 (15.9%)  Commented: 4 (1.1%)  Filled: 244 (68.2%)  Empty: 114 (31.8%)  Max Logic Depth: 3
