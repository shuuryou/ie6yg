VERSION 5.00
Begin VB.Form frmNetworkPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Network Password Form Layout"
   ClientHeight    =   3405
   ClientLeft      =   7350
   ClientTop       =   4920
   ClientWidth     =   5910
   Icon            =   "frmNetworkPassword.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "__"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "__"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2070
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   2010
      Width           =   3165
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   2070
      TabIndex        =   6
      Top             =   1560
      Width           =   3165
   End
   Begin VB.CheckBox chkSave 
      Caption         =   "__"
      Enabled         =   0   'False
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   2415
      Width           =   3735
   End
   Begin VB.Label lblRealm 
      Caption         =   "__Realm__"
      Height          =   255
      Left            =   2070
      TabIndex        =   4
      Top             =   1170
      Width           =   3165
   End
   Begin VB.Label lblSite 
      Caption         =   "__Server__"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2070
      TabIndex        =   2
      Top             =   735
      Width           =   3165
   End
   Begin VB.Label lblPassword 
      Caption         =   "__"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   2010
      Width           =   975
   End
   Begin VB.Label lblUsername 
      Caption         =   "__"
      Height          =   255
      Left            =   945
      TabIndex        =   5
      Top             =   1605
      Width           =   975
   End
   Begin VB.Label lblRealmDesc 
      Caption         =   "__"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1170
      Width           =   975
   End
   Begin VB.Label lblSiteDesc 
      Caption         =   "__"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   735
      Width           =   975
   End
   Begin VB.Label lblAction 
      Caption         =   "__"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   360
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmNetworkPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' It would be nice not to have to do all of this work and just
' rely on InternetErrorDlg(), but hRequest requires a HINTERNET
' handle that was passed to HttpSendRequest() before. Since all
' of the authentication handling is inside CEFsharp, that's
' impossible and just passing a NULL handle doesn't work.

' This form layout is incorrect in WinXP and later, since they
' use CredUIPromptForCredentials. I don't care at the moment.

' Using IE's native (and insecure on Win9x) credential storage
' is extremely difficult and the required APIs are not really
' documented. Most source code on the web implements reading it,
' and I didn't want to waste days or weeks trying to implement
' writing to it. Hence rolling my own password saving feature.

' I don't consider it secure; it uses SHA1 to derive a key and
' RC4 to encrypt, both of which have known problems, and yet here
' we are, because Win9x doesn't support any modern cryptographic
' algorithms.

' Please consider the password security implemented here to be
' one level above XOR.

' Data is stored in the registry at:
' HKEY_CURRENT_USER\Software\IE6YG\Password List
' Value <SiteHash>_U = Encrypted ANSI username bytes
' Value <SiteHash>_P = Encrypted ANSI password bytes

' <SiteHash> is the SHA1 hash of (sitename + server hash +
' client hash) in hex, but note that hash bytes below 0x10 are
' not padded with a 0.

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByRef pszContainer As Any, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal AlgID As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As Long, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As Long, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal AlgID As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As Long, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As Long, ByRef pdwDataLen As Long) As Long

Private Const MS_DEF_PROV As String = "Microsoft Base Cryptographic Provider v1.0"

Private Const PROV_RSA_FULL As Long = &H1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Private Const ALG_CLASS_HASH As Long = &H8000& ' & at the end is important
Private Const ALG_CLASS_DATA_ENCRYPT As Long = &H6000& ' & at the end is important

Private Const ALG_TYPE_STREAM As Long = &H800& ' & at the end is important
Private Const ALG_TYPE_ANY As Long = &H0

Private Const ALG_SID_SHA1 As Long = &H4
Private Const ALG_SID_RC4 As Long = &H1

Private Const CRYPT_NO_SALT = &H10

Private Const HP_HASHVAL As Long = &H2
Private Const HP_HASHSIZE As Long = &H4

Private m_hWndParent As Long
Private m_hIconDlgIcon As Long

Private m_ServerHostID As String
Private m_ClientHostID As String

Private m_EnteredUsername As String
Private m_EnteredPassword As String

Private m_Result As VbMsgBoxResult

Private Sub cmdCancel_Click()

    m_Result = vbCancel
    Unload Me

End Sub

Private Sub cmdOK_Click()

    m_Result = vbOK

    If chkSave.value = vbChecked And (txtUsername.Text <> "" Or txtPassword.Text <> "") Then
        StoreCredentialsInRegistry
    End If

    m_EnteredUsername = txtUsername.Text
    m_EnteredPassword = txtPassword.Text

    Unload Me

End Sub

Private Function DecryptString(enc() As Byte, pwd As String) As String

  Dim hProvider As Long
  Dim hHash As Long
  Dim hKey As Long

  Dim bytData() As Byte
  Dim lngDataLen As Long

    If CryptAcquireContext(hProvider, ByVal 0, MS_DEF_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        modLogging.WriteLineToLog "DecryptString: Failed to get handle for MS_DEF_PROV context. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    If CryptCreateHash(hProvider, (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1), 0, 0, hHash) = 0 Then
        modLogging.WriteLineToLog "DecryptString: Failed to get handle for hash algorithm. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    If CryptHashData(hHash, StrPtr(pwd), LenB(pwd), 0) = 0 Then
        modLogging.WriteLineToLog "DecryptString: Failed to derive hash from password. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    If CryptDeriveKey(hProvider, (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4), hHash, CRYPT_NO_SALT, hKey) = 0 Then
        modLogging.WriteLineToLog "DecryptString: Failed to derive key from password hash. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    bytData = enc
    lngDataLen = UBound(bytData)

    If CryptDecrypt(hKey, 0, 1, 0, VarPtr(bytData(0)), lngDataLen) = 0 Then
        modLogging.WriteLineToLog "DecryptString: Failed to decrypt. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    ReDim Preserve bytData(lngDataLen - 1)

    DecryptString = StrConv(bytData, vbUnicode)

    Erase bytData

done:

    If hKey <> 0 Then
        CryptDestroyKey hKey
    End If

    If hHash <> 0 Then
        CryptDestroyHash hHash
    End If

    If hProvider <> 0 Then
        CryptReleaseContext hProvider, 0
    End If

End Function

Public Property Get DialogResult() As VbMsgBoxResult

    DialogResult = m_Result

End Property

Private Function EncryptString(str As String, pwd As String) As Byte()

  Dim hProvider As Long
  Dim hHash As Long
  Dim hKey As Long
  Dim cbCipher As Long

  Dim bytData() As Byte
  Dim lngDataLen As Long

    If CryptAcquireContext(hProvider, ByVal 0, MS_DEF_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        modLogging.WriteLineToLog "EncryptString: Failed to get handle for MS_DEF_PROV context. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    If CryptCreateHash(hProvider, (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1), 0, 0, hHash) = 0 Then
        modLogging.WriteLineToLog "EncryptString: Failed to get handle for hash algorithm. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    If CryptHashData(hHash, StrPtr(pwd), LenB(pwd), 0) = 0 Then
        modLogging.WriteLineToLog "EncryptString: Failed to derive hash from password. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    If CryptDeriveKey(hProvider, (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4), hHash, CRYPT_NO_SALT, hKey) = 0 Then
        modLogging.WriteLineToLog "EncryptString: Failed to derive key from password hash. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    bytData = StrConv(str, vbFromUnicode)
    lngDataLen = UBound(bytData) + 1
    cbCipher = lngDataLen

    If CryptEncrypt(hKey, 0, 1, 0, 0, cbCipher, 0) = 0 Then
        modLogging.WriteLineToLog "EncryptString: Failed to determine length of encrypted data. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    ReDim Preserve bytData(cbCipher + 1)

    If CryptEncrypt(hKey, 0, 1, 0, VarPtr(bytData(0)), lngDataLen, cbCipher) = 0 Then
        modLogging.WriteLineToLog "EncryptString: Failed to encrypt. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    EncryptString = bytData

done:

    If hKey <> 0 Then
        CryptDestroyKey hKey
    End If

    If hHash <> 0 Then
        CryptDestroyHash hHash
    End If

    If hProvider <> 0 Then
        CryptReleaseContext hProvider, 0
    End If

End Function

Private Sub Form_Initialize()

    m_hWndParent = -1

    Me.Caption = LoadResString(400) ' Enter Network Password

    lblAction.Caption = LoadResString(401) ' Please type your user name and password.
    lblSiteDesc.Caption = LoadResString(402) ' Site
    lblRealmDesc.Caption = LoadResString(403) ' Realm
    lblUsername.Caption = LoadResString(404) ' User Name
    lblPassword.Caption = LoadResString(405) ' Password

    chkSave.Caption = LoadResString(406) ' Save this password in your password list

    cmdOK.Caption = LoadResString(407) ' OK
    cmdCancel.Caption = LoadResString(408) ' Cancel

    m_hIconDlgIcon = LoadImage(App.hInstance, MAKEINTRESOURCE(400), IMAGE_ICON, 32, 32, 0&)

    imgIcon.Picture = HandleToPicture(m_hIconDlgIcon, vbPicTypeIcon)

End Sub

Private Sub Form_Load()

    If m_hWndParent <> -1 Then
        SetParent Me.hWnd, m_hWndParent
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    DestroyIcon m_hIconDlgIcon

End Sub

Public Sub GetCredentialsFromRegistry()

  Dim strSiteHash As String
  Dim strUsername As String
  Dim strPassword As String

  Dim objRegistry As New CRegistry

    strSiteHash = HashString(m_ServerHostID & m_ClientHostID & lblSite.Caption)

    With objRegistry
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\IE6YG\Password List"

        If Not .KeyExists Then
            modLogging.WriteLineToLog "GetCredentialsFromRegistry: No credentials in registry."
            Exit Sub
        End If

        .Default = -1
        .ValueType = REG_BINARY
        .ValueKey = strSiteHash & "_U"

        If VarType(.value) <> (vbByte Or vbArray) Then
            modLogging.WriteLineToLog "GetCredentialsFromRegistry: SiteHash_U not found."
            Exit Sub
        End If

        strUsername = DecryptString(.value, m_ServerHostID & m_ClientHostID)

        .ValueKey = strSiteHash & "_P"

        If VarType(.value) <> (vbByte Or vbArray) Then
            modLogging.WriteLineToLog "GetCredentialsFromRegistry: SiteHash_P not found."
            Exit Sub
        End If

        strPassword = DecryptString(.value, m_ServerHostID & m_ClientHostID)
    End With

    Set objRegistry = Nothing

    txtUsername.Text = strUsername
    txtPassword.Text = strPassword

    If txtUsername.Text <> "" Or txtPassword.Text <> "" Then
        chkSave.value = vbChecked
    End If

End Sub

Private Function HashString(str As String) As String

  Dim hProvider As Long
  Dim hHash As Long
  Dim hKey As Long

  Dim bytData() As Byte
  Dim lngDataLen As Long

  Dim strRet As String

    If CryptAcquireContext(hProvider, ByVal 0, MS_DEF_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        modLogging.WriteLineToLog "HashString: Failed to get handle for MS_DEF_PROV context. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    If CryptCreateHash(hProvider, (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1), 0, 0, hHash) = 0 Then
        modLogging.WriteLineToLog "HashString: Failed to get handle for hash algorithm. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    If CryptHashData(hHash, StrPtr(str), LenB(str), 0) = 0 Then
        modLogging.WriteLineToLog "HashString: Failed to derive hash from password. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    If CryptGetHashParam(hHash, HP_HASHVAL, 0, lngDataLen, 0) = 0 Then
        modLogging.WriteLineToLog "HashString: Failed to get size of hash value. HRESULT=" & Err.LastDllError
        GoTo done
    End If

    ReDim bytData(lngDataLen - 1)

    If CryptGetHashParam(hHash, HP_HASHVAL, VarPtr(bytData(0)), lngDataLen, 0) = 0 Then
        modLogging.WriteLineToLog "HashString: Failed to read hash value. HRESULT=" & Err.LastDllError
        GoTo done

    End If

  Dim i As Long
    For i = 0 To lngDataLen - 1
        strRet = strRet & Hex$(bytData(i))
    Next i

    HashString = strRet

done:

    If hHash <> 0 Then
        CryptDestroyHash hHash
    End If

    If hProvider <> 0 Then
        CryptReleaseContext hProvider, 0
    End If

End Function

Public Property Let ParentWindowHandle(hWnd As Long)

    m_hWndParent = hWnd

End Property

Public Property Get Password() As String

    Password = m_EnteredPassword

End Property

Public Property Let Realm(strRealm As String)

    lblRealm.Caption = strRealm

End Property

Public Sub SetHostID(strServer As String, strClient As String)

    m_ServerHostID = strServer
    m_ClientHostID = strClient
    chkSave.Enabled = (m_ServerHostID <> "" And m_ClientHostID <> "")

End Sub

Public Property Let Site(strSite As String)

    lblSite.Caption = strSite

End Property

Private Sub StoreCredentialsInRegistry()

  Dim bytUsername() As Byte
  Dim bytPassword() As Byte

  Dim strSiteHash As String

  Dim objRegistry As New CRegistry

    bytUsername = EncryptString(txtUsername.Text, m_ServerHostID & m_ClientHostID)

    If ((Not bytUsername) = -1) Then
        MsgBox LoadResString(410), vbCritical, Me.Caption
        Exit Sub
    End If

    bytPassword = EncryptString(txtPassword.Text, m_ServerHostID & m_ClientHostID)

    If ((Not bytPassword) = -1) Then
        MsgBox LoadResString(410), vbCritical, Me.Caption
        Exit Sub
    End If

    strSiteHash = HashString(m_ServerHostID & m_ClientHostID & lblSite.Caption)

    With objRegistry
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\IE6YG\Password List"
        .ValueType = REG_BINARY
        .ValueKey = strSiteHash & "_U"
        .value = bytUsername()

        Erase bytUsername

        .ValueKey = strSiteHash & "_P"
        .value = bytPassword()

        Erase bytPassword
    End With

    Set objRegistry = Nothing

End Sub

Public Property Let URL(strUrl As String)

    lblSite.ToolTipText = strUrl

End Property

Public Property Get Username() As String

    Username = m_EnteredUsername

End Property

':) Ulli's VB Code Formatter V2.24.17 (2022-Dec-06 07:14)  Decl: 74  Code: 388  Total: 462 Lines
':) CommentOnly: 27 (5.8%)  Commented: 12 (2.6%)  Filled: 312 (67.5%)  Empty: 150 (32.5%)  Max Logic Depth: 3
