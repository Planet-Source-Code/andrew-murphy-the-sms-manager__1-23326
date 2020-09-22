VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SMS Manager"
   ClientHeight    =   1515
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "&Internet Browser"
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send/Get &Emails"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send An &SMS"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmSMS.Show
Unload Me
End Sub

Private Sub Command2_Click()
TestMailLite.Show
Unload Me
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
frmBrowser.Show
Unload Me
End Sub
