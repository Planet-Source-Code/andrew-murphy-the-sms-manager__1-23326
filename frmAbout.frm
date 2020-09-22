VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   1290
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4350
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   890.381
   ScaleMode       =   0  'User
   ScaleWidth      =   4084.875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "You can send SMS or Text messages to any phone with this program and whats more it is 100% FREE"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  frmSMS.Show
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub
