VERSION 5.00
Object = "{6395F295-D138-11D1-A6B4-00AA002075DA}#1.0#0"; "MAILLITE.OCX"
Begin VB.Form TestMailLite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test MailLite"
   ClientHeight    =   6930
   ClientLeft      =   2490
   ClientTop       =   1155
   ClientWidth     =   7335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "TestMailLite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5520
      TabIndex        =   34
      Top             =   2520
      Width           =   1695
   End
   Begin MailLitePrj.MailLite MailLite1 
      Left            =   5280
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      UserHeaders     =   ""
      SenderEmail     =   ""
      RecipientEmail  =   ""
      CarbonCopyEmail =   ""
      EmailSubject    =   ""
      EmailBody       =   ""
      HostNameSmtp    =   ""
      HostNamePop3    =   ""
      SenderName      =   ""
      RecipientName   =   ""
      CarbonCopyName  =   ""
      UserAccount     =   ""
      UserPassword    =   ""
      EmailPath       =   ""
      EmailFilter     =   ""
   End
   Begin VB.TextBox EmailFilter 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   960
      TabIndex        =   3
      Top             =   1380
      Width           =   3075
   End
   Begin VB.TextBox EmailUIDL 
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2460
      Width           =   4395
   End
   Begin VB.TextBox CarbonCopyEmail 
      Height          =   315
      Left            =   4140
      TabIndex        =   13
      Top             =   3960
      Width           =   3075
   End
   Begin VB.TextBox CarbonCopyName 
      Height          =   315
      Left            =   960
      TabIndex        =   12
      Top             =   3960
      Width           =   3075
   End
   Begin VB.FileListBox InBox 
      Height          =   870
      Left            =   4200
      Pattern         =   "*.mbx"
      TabIndex        =   29
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox EmailDate 
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   4395
   End
   Begin VB.TextBox RecipientName 
      Height          =   315
      Left            =   960
      TabIndex        =   10
      Top             =   3540
      Width           =   3075
   End
   Begin VB.CheckBox HostDelete 
      Caption         =   "Delete email from host"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   180
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox Status 
      Height          =   735
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6120
      Width           =   6255
   End
   Begin VB.TextBox SenderName 
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Top             =   3120
      Width           =   3075
   End
   Begin VB.TextBox UserPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   3075
   End
   Begin VB.TextBox UserAccount 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   540
      Width           =   3075
   End
   Begin VB.TextBox HostName 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3075
   End
   Begin VB.CommandButton Send 
      Caption         =   "Send Emails"
      Height          =   495
      Left            =   5520
      TabIndex        =   17
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox EmailBody 
      Height          =   915
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   4800
      Width           =   6255
   End
   Begin VB.TextBox EmailSubject 
      Height          =   315
      Left            =   960
      TabIndex        =   14
      Top             =   4380
      Width           =   6255
   End
   Begin VB.TextBox RecipientEmail 
      Height          =   315
      Left            =   4140
      TabIndex        =   11
      Top             =   3540
      Width           =   3075
   End
   Begin VB.TextBox SenderEmail 
      Height          =   315
      Left            =   4140
      TabIndex        =   9
      Top             =   3120
      Width           =   3075
   End
   Begin VB.CommandButton Check 
      Caption         =   "Check Emails"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "Filter:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label Label14 
      Caption         =   "UIDL:"
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   2520
      Width           =   795
   End
   Begin VB.Label Label13 
      Caption         =   "Cc:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   4020
      Width           =   795
   End
   Begin VB.Label Label12 
      Caption         =   "InBox"
      Height          =   195
      Left            =   4200
      TabIndex        =   30
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2100
      Width           =   795
   End
   Begin VB.Label Label10 
      Caption         =   "Email address:"
      Height          =   195
      Left            =   4140
      TabIndex        =   27
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7200
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label9 
      Caption         =   "Status:"
      Height          =   315
      Left            =   120
      TabIndex        =   26
      Top             =   6120
      Width           =   795
   End
   Begin VB.Label Label8 
      Caption         =   "Full Name"
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7200
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label Label7 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1020
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "Account:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label5 
      Caption         =   "Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   180
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4860
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3180
      Width           =   795
   End
End
Attribute VB_Name = "TestMailLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Public CheckOk As Boolean
Public Busy As Integer
Public Msg As String

Private Sub Command1_Click()
End
End Sub

Private Sub UserAccount_GotFocus()
    UserAccount.SelStart = 0
    UserAccount.SelLength = Len(UserAccount)
End Sub

Private Sub CarbonCopyEmail_GotFocus()
    CarbonCopyEmail.SelStart = 0
    CarbonCopyEmail.SelLength = Len(CarbonCopyEmail)
End Sub

Private Sub CarbonCopyName_GotFocus()
    CarbonCopyName.SelStart = 0
    CarbonCopyName.SelLength = Len(CarbonCopyName)
End Sub

Private Sub EmailFilter_GotFocus()
    EmailFilter.SelStart = 0
    EmailFilter.SelLength = Len(EmailFilter)
End Sub

Private Sub Form_Load()
    MailLite1.EmailPath = App.Path
    InBox.Path = MailLite1.EmailPath
End Sub

Private Sub HostName_GotFocus()
    HostName.SelStart = 0
    HostName.SelLength = Len(HostName)
End Sub

Private Sub InBox_DblClick()
    If MailLite1.LoadEmail(InBox.FileName) Then
        UpdateTextBoxes
    End If
End Sub

Private Sub UpdateTextBoxes()
    EmailDate = MailLite1.EmailDate
    EmailUIDL = MailLite1.EmailUIDL
    SenderName = MailLite1.SenderName
    SenderEmail = MailLite1.SenderEmail
    RecipientName = MailLite1.RecipientName
    RecipientEmail = MailLite1.RecipientEmail
    CarbonCopyName = MailLite1.CarbonCopyName
    CarbonCopyEmail = MailLite1.CarbonCopyEmail
    EmailSubject = MailLite1.EmailSubject
    EmailBody = MailLite1.EmailBody
End Sub

Private Sub MailLite1_Status(ByVal Code As Integer, ByVal Description As String)
    Status.SelStart = Len(Status.Text)
    Status.SelText = vbCrLf & CStr(Code) & " - " & Description
    If Len(Status.Text) > 32000 Then
        Status.Text = Mid(Status.Text, Len(Status.Text) - 32000)
    End If
    If Code = 117 Then
        InBox.Refresh
        UpdateTextBoxes
    End If
End Sub

Private Sub Check_Click()
    If Busy Then Exit Sub
    Busy = True
    
    MailLite1.HostPortPop3 = 110                        'default
    MailLite1.HostNamePop3 = HostName
    MailLite1.HostDelete = IIf(HostDelete, True, False)
    MailLite1.UserAccount = UserAccount
    MailLite1.UserPassword = UserPassword
    MailLite1.EmailFilter = EmailFilter
    'MailLite1.EmailFilter = "X-Priority: 1"
    'MailLite1.EmailPath = App.Path
    
    CheckOk = MailLite1.CheckEmail(True, True)
    If Not CheckOk Then
        If MailLite1.TotalEmails = MailLite1.TotalChecked Then
            'it is a minor bug, all the mail was retrieved ok.
            CheckOk = True
        Else
            If MailLite1.TotalChecked Then
                Msg = "We can read only " & CStr(MailLite1.TotalChecked) & " emails."
            Else
                Msg = "We can not read any email."
            End If
            MsgBox "Some problem reading emails! - " & Msg & " - " & _
                    "Error " & MailLite1.Error.Number & " - " & _
                    MailLite1.Error.Description & ".", _
                    vbExclamation, Me.Caption
        End If
    End If
    If CheckOk Then
        If MailLite1.TotalChecked > 1 Then
            MsgBox "Was checked successfully " & CStr(MailLite1.TotalChecked) & " emails!", vbInformation, Me.Caption
        ElseIf MailLite1.TotalChecked = 1 Then
            MsgBox "One email was checked successfully!", vbInformation, Me.Caption
        Else
            MsgBox "There are not any email to retrieve from the mailbox!", vbInformation, Me.Caption
        End If
    End If
    Busy = False
End Sub

Private Sub EmailBody_GotFocus()
    EmailBody.SelStart = 0
    EmailBody.SelLength = Len(EmailBody)
End Sub

Private Sub UserPassword_GotFocus()
    UserPassword.SelStart = 0
    UserPassword.SelLength = Len(UserPassword)
End Sub

Private Sub RecipientEmail_GotFocus()
    RecipientEmail.SelStart = 0
    RecipientEmail.SelLength = Len(RecipientEmail)
End Sub

Private Sub RecipientName_GotFocus()
    RecipientName.SelStart = 0
    RecipientName.SelLength = Len(RecipientName)
End Sub

Private Sub Send_Click()
    If Busy Then Exit Sub
    Busy = True
    MailLite1.HostPortSmtp = 25                         'default
    MailLite1.HostNameSmtp = HostName
    MailLite1.SenderName = SenderName                   'optional
    MailLite1.SenderEmail = SenderEmail
    MailLite1.RecipientName = RecipientName             'optional
    MailLite1.RecipientEmail = RecipientEmail
    MailLite1.CarbonCopyName = CarbonCopyName           'optional
    MailLite1.CarbonCopyEmail = CarbonCopyEmail         'optional
    MailLite1.EmailSubject = EmailSubject
    MailLite1.UserHeaders = "X-Priority: 1" & vbCrLf    'optional
    MailLite1.EmailBody = EmailBody
    
    If MailLite1.SendEmail() Then
        MsgBox "The email was successfully sended!", vbInformation, Me.Caption
    Else
        MsgBox "Some problem sending emails! - " & _
                "Error " & MailLite1.Error.Number & " - " & _
                MailLite1.Error.Description & ".", _
                vbExclamation, Me.Caption
    End If
    Busy = False
End Sub

Private Sub SenderEmail_GotFocus()
    SenderEmail.SelStart = 0
    SenderEmail.SelLength = Len(SenderEmail)
End Sub

Private Sub SenderName_GotFocus()
    SenderName.SelStart = 0
    SenderName.SelLength = Len(SenderName)
End Sub

Private Sub EmailSubject_GotFocus()
    EmailSubject.SelStart = 0
    EmailSubject.SelLength = Len(EmailSubject)
End Sub
