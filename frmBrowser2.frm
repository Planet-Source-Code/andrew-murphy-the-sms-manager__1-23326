VERSION 5.00
Object = "{B57329AE-4D0B-11D5-BFD2-EB69A187A478}#1.0#0"; "THEWEB.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "Web Browser"
   ClientHeight    =   7275
   ClientLeft      =   1110
   ClientTop       =   1275
   ClientWidth     =   9720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   9720
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin browser.Webbrowser Webbrowser1 
      Height          =   5295
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9340
      Source          =   $"frmBrowser2.frx":0000
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   255
      Left            =   8280
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Spanish"
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Decrease"
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Increase"
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go !"
      Height          =   255
      Left            =   9100
      TabIndex        =   4
      Top             =   1500
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Source"
      Height          =   615
      Left            =   6960
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   9615
   End
   Begin VB.Image Image4 
      Height          =   1425
      Index           =   2
      Left            =   8280
      Picture         =   "frmBrowser2.frx":0020
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image Image4 
      Height          =   1425
      Index           =   1
      Left            =   5760
      Picture         =   "frmBrowser2.frx":0119
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image Image4 
      Height          =   1425
      Index           =   0
      Left            =   7080
      Picture         =   "frmBrowser2.frx":0212
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image Image3 
      Height          =   1425
      Left            =   2640
      Picture         =   "frmBrowser2.frx":030B
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   1425
      Left            =   3480
      Picture         =   "frmBrowser2.frx":13C1
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   4320
      Picture         =   "frmBrowser2.frx":24B7
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image Image5 
      Height          =   1425
      Left            =   1800
      Picture         =   "frmBrowser2.frx":3BB3
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image7 
      Height          =   1425
      Left            =   960
      Picture         =   "frmBrowser2.frx":4CA6
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image6 
      Height          =   1425
      Left            =   0
      Picture         =   "frmBrowser2.frx":5CBC
      Top             =   0
      Width           =   1005
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: U.Sudeep Nayak
'Yes, it's here, with a lot of bugs for you to report.
'New Feature: It's FAAAAASSSSSTER ! Yeah! It's at least
'200 % Faster than any other browser i've seen.
''''''''''''''
'You can find & buy these images at www.guistuff.com at "Odd Browser"
'There are a lot more things you can do with the browser, not shown here.
'like webbrowser.translate, webbrowser.increasetextsize ,
'webbrowser1.GoNews, webbrowser1.LocalIP,
'webbrowser1.LocationURL, webbrowser1.Offline,
'webbrowser1.Pausenav (pauses navigation)
'webbrowser1.Resumenav(Resumes navigation)
'webbrowser1.Refresh,webbrowser1.Source,etc.
'More to come next version, including better right-click
'menu. Only your suggestions would help.
'-------------------------------------------Sudeep.

Private Sub Combo1_KeyPress(KeyAscii As Integer)
'Author: U.Sudeep Nayak
If KeyAscii = 13 Then Webbrowser1.Navigate (Combo1.Text)
End Sub

Private Sub Command1_Click()
MsgBox (Webbrowser1.Source)

End Sub
'Author: U.Sudeep Nayak

Private Sub Command2_Click()
Webbrowser1.Navigate (Combo1.Text)
End Sub

Private Sub Command3_Click()
Webbrowser1.IncreaseFontSize (1)
End Sub

Private Sub Command4_Click()
Webbrowser1.DecreaseFontSize (1)
End Sub

Private Sub Command5_Click()
Webbrowser1.Translate "English", "Spanish"
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Image1_Click()
MsgBox ("C'mon , Use your imagination for this will ya?")
End Sub

Private Sub Image2_Click()
search$ = InputBox("Search For What", "Search")
Webbrowser1.GoSearch (search$)
End Sub

Private Sub Image3_Click()
Webbrowser1.GoHome
End Sub


Private Sub Image5_Click()
Webbrowser1.Stopnav

End Sub

Private Sub Image6_Click()
Webbrowser1.GoBack

'Author: U.Sudeep Nayak
End Sub

Private Sub Image7_Click()
Webbrowser1.GoForward

End Sub

Private Sub Webbrowser1_StatusTextChange(ByVal Text As String)
Label1.Caption = Text
End Sub

'Author: U.Sudeep Nayak

