VERSION 5.00
Begin VB.Form frmCertificateError 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Certificate Error Form Layout"
   ClientHeight    =   4695
   ClientLeft      =   7425
   ClientTop       =   3870
   ClientWidth     =   5670
   Icon            =   "frmCertificateError.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdViewCert 
      Caption         =   "__"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "__"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "__"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lblItem 
      Caption         =   "__"
      Height          =   495
      Index           =   3
      Left            =   1200
      TabIndex        =   4
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label lblItem 
      Caption         =   "__"
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label lblItem 
      Caption         =   "__"
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label lblItem 
      Caption         =   "__"
      Height          =   855
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label lblProceed 
      Caption         =   "__"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Image imgItem 
      Height          =   240
      Index           =   3
      Left            =   840
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgItem 
      Height          =   240
      Index           =   2
      Left            =   840
      Top             =   2520
      Width           =   240
   End
   Begin VB.Image imgItem 
      Height          =   240
      Index           =   1
      Left            =   840
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image imgItem 
      Height          =   240
      Index           =   0
      Left            =   840
      Top             =   960
      Width           =   240
   End
   Begin VB.Label lblMessage 
      Caption         =   "__"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmCertificateError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' It would be nice not to have to do all of this work and just
' rely on InternetErrorDlg(), but hRequest requires a HINTERNET
' handle that was passed to HttpSendRequest() before. Since all
' of the certificate handling stuff comes from CEFsharp, that's
' impossible and just passing a NULL handle doesn't work.

Public Enum CertificateStates
    None = 0
    OverallOK = &H1
    OverallBad = &H2
    UntrustedIssuer = &H4
    TrustedIssuer = &H8
    DateValid = &H10
    DateInvalid = &H20
    NameValid = &H40
    NameInvalid = &H80
    StrongCert = &H100
    WeakCert = &H200
    Revoked = &H400
    DateInvalidTooLong = &H800
    ChromeCTFail = &H1000
End Enum
#If False Then
Private None, OverallOK, OverallBad, UntrustedIssuer, TrustedIssuer, DateValid, DateInvalid, NameValid, NameInvalid, StrongCert, WeakCert, _
        Revoked, DateInvalidTooLong, ChromeCTFail
#End If

Private m_hIconOverallOK As Long
Private m_hIconOverallBad As Long
Private m_hIconValid As Long
Private m_hIconInvalid As Long

Private m_hWndParent As Long

Private m_CertFile As String

Private m_Result As VbMsgBoxResult

Public Property Let CertificateFile(path As String)

    GetAttr path ' Will raise an error if it doesn't exist.

    m_CertFile = path

End Property

Public Property Let CertificateState(state As CertificateStates)

    If (state And OverallOK) <> 0 Then
        Me.Caption = LoadResString(222) ' Security Information
        lblMessage.Caption = LoadResString(201)
        imgIcon.Picture = HandleToPicture(m_hIconOverallOK, vbPicTypeIcon)
      ElseIf (state And OverallBad) <> 0 Then
        Me.Caption = LoadResString(215) ' Security Alert
        lblMessage.Caption = LoadResString(200)
        imgIcon.Picture = HandleToPicture(m_hIconOverallBad, vbPicTypeIcon)
    End If

    If (state And TrustedIssuer) <> 0 Then
        lblItem(0).Caption = LoadResString(203)
        imgItem(0).Picture = HandleToPicture(m_hIconValid, vbPicTypeIcon)
      ElseIf (state And Revoked) <> 0 Then
        lblItem(0).Caption = LoadResString(210)
        imgItem(0).Picture = HandleToPicture(m_hIconInvalid, vbPicTypeIcon)
      ElseIf (state And UntrustedIssuer) <> 0 Then
        lblItem(0).Caption = LoadResString(202)
        imgItem(0).Picture = HandleToPicture(m_hIconInvalid, vbPicTypeIcon)
      ElseIf (state And ChromeCTFail) <> 0 Then
        lblItem(0).Caption = LoadResString(217)
        imgItem(0).Picture = HandleToPicture(m_hIconInvalid, vbPicTypeIcon)
    End If

    If (state And DateValid) <> 0 Then
        lblItem(1).Caption = LoadResString(204)
        imgItem(1).Picture = HandleToPicture(m_hIconValid, vbPicTypeIcon)
      ElseIf (state And DateInvalid) <> 0 Then
        lblItem(1).Caption = LoadResString(205)
        imgItem(1).Picture = HandleToPicture(m_hIconInvalid, vbPicTypeIcon)
      ElseIf (state And DateInvalidTooLong) <> 0 Then
        lblItem(1).Caption = LoadResString(216)
        imgItem(1).Picture = HandleToPicture(m_hIconInvalid, vbPicTypeIcon)
    End If

    If (state And NameValid) <> 0 Then
        lblItem(2).Caption = LoadResString(207)
        imgItem(2).Picture = HandleToPicture(m_hIconValid, vbPicTypeIcon)
      ElseIf (state And NameInvalid) <> 0 Then
        lblItem(2).Caption = LoadResString(206)
        imgItem(2).Picture = HandleToPicture(m_hIconInvalid, vbPicTypeIcon)
    End If

    If (state And StrongCert) <> 0 Then
        lblItem(3).Caption = LoadResString(208)
        imgItem(3).Picture = HandleToPicture(m_hIconValid, vbPicTypeIcon)
      ElseIf (state And WeakCert) <> 0 Then
        lblItem(3).Caption = LoadResString(209)
        imgItem(3).Picture = HandleToPicture(m_hIconInvalid, vbPicTypeIcon)
    End If

End Property

Private Sub cmdNo_Click()

    m_Result = vbNo
    Unload Me

End Sub

Private Sub cmdViewCert_Click()

    On Error GoTo EH
    If m_CertFile = "" Then
        Err.Raise -1, "cmdViewCert", "No certificate file."
    End If

    ' m_CertFile is a short file name
    Shell "rundll32.exe cryptext.dll,CryptExtOpenCER " & m_CertFile, vbNormalFocus

Exit Sub

EH:
    modLogging.WriteLineToLog "cmdViewCert: Error: " & Err.Number & " (" & Err.Description & ")"
    MsgBox LoadResString(218), vbCritical Or vbOKOnly, Me.Caption ' The certificate could not be opened.

End Sub

Private Sub cmdYes_Click()

    m_Result = vbYes
    Unload Me

End Sub

Private Sub Form_Initialize()

    m_hWndParent = -1

    lblProceed.Caption = LoadResString(211) ' Do you want to proceed?

    cmdYes.Caption = LoadResString(212) ' Yes
    cmdNo.Caption = LoadResString(213) ' No
    cmdViewCert.Caption = LoadResString(214) ' View Certificate

    m_hIconOverallBad = LoadImage(App.hInstance, MAKEINTRESOURCE(201), IMAGE_ICON, 32, 32, 0&)
    m_hIconOverallOK = LoadImage(App.hInstance, MAKEINTRESOURCE(202), IMAGE_ICON, 32, 32, 0&)
    m_hIconInvalid = LoadImage(App.hInstance, MAKEINTRESOURCE(203), IMAGE_ICON, 16, 16, 0&)
    m_hIconValid = LoadImage(App.hInstance, MAKEINTRESOURCE(204), IMAGE_ICON, 16, 16, 0&)

End Sub

Private Sub Form_Load()

    If m_hWndParent <> -1 Then
        SetParent Me.hWnd, m_hWndParent
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    DestroyIcon m_hIconOverallBad
    DestroyIcon m_hIconOverallOK
    DestroyIcon m_hIconInvalid
    DestroyIcon m_hIconValid

End Sub

Public Property Let NoPromptMode(enable As Boolean)

  ' IE6 just shows the certificate details screen (that is lazily
  ' opened here using RUNDLL32 to avoid translating most of the
  ' WINCRYPT.H to VB6. Since root certificates on ancient Win9x
  ' are terribly out of date, the warning dialog is repurposed
  ' to at least show a summary of what Chromium thinks about the
  ' certificate.

    If enable Then
        cmdNo.Visible = False
        cmdNo.Cancel = False
        cmdNo.Default = False

        cmdYes.Caption = LoadResString(221)
        cmdYes.Cancel = True
        cmdYes.Default = True

        lblProceed.Visible = False
      Else
        cmdYes.Caption = LoadResString(212)
        cmdYes.Cancel = False
        cmdYes.Default = False

        cmdNo.Visible = False
        cmdNo.Cancel = True
        cmdNo.Default = True

        lblProceed.Visible = True
    End If

End Property

Public Property Let ParentWindowHandle(hWnd As Long)

    m_hWndParent = hWnd

End Property

Public Property Get DialogResult() As VbMsgBoxResult

    DialogResult = m_Result

End Property

':) Ulli's VB Code Formatter V2.24.17 (2022-Nov-18 00:20)  Decl: 41  Code: 206  Total: 247 Lines
':) CommentOnly: 14 (5.7%)  Commented: 8 (3.2%)  Filled: 182 (73.7%)  Empty: 65 (26.3%)  Max Logic Depth: 3
