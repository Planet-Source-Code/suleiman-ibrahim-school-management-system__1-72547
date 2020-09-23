VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404000&
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12690
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   12690
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
      Begin VB.Label Label1 
         BackColor       =   &H00404000&
         Caption         =   "SUB MENU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.ListBox lstLoad_Forms 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3840
      ItemData        =   "frmMain.frx":0442
      Left            =   840
      List            =   "frmMain.frx":0444
      TabIndex        =   3
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   4335
      Left            =   4440
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblIBRUQ_Tech 
      BackColor       =   &H00404000&
      Caption         =   "IBRUQ COMPUTER INSTITUTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   720
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
      Caption         =   "LOKOJA KOGI STATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   1800
      Width           =   5415
   End
   Begin VB.Label lblDivision 
      BackColor       =   &H00404000&
      Caption         =   "A DIVISION OF IBRUQ TECH NIG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   3900
      Left            =   4920
      Picture         =   "frmMain.frx":0446
      Top             =   2760
      Width           =   6720
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      WindowList      =   -1  'True
      Begin VB.Menu mnuLoin 
         Caption         =   "Login Staff"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log Out Staff"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmMain.Caption = "Main Menu"
lstLoad_Forms.AddItem "Registration"
lstLoad_Forms.AddItem "Student Result"
lstLoad_Forms.AddItem "Fee Payment"
lstLoad_Forms.AddItem "Print Situation Report"
lstLoad_Forms.AddItem "About IBRUQ-TECH NIG"
lstLoad_Forms.AddItem "Exit"
lstLoad_Forms.Enabled = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Exit Program?", vbYesNo + vbQuestion, "Exit") = vbYes Then
End
Else
Cancel = True
End If
End Sub
Private Sub lstLoad_Forms_Click()
With lstLoad_Forms
If lstLoad_Forms.Text = "Registration" Then
frmRegist.Show vbModal
End If
If lstLoad_Forms.Text = "Student Result" Then
frmResult.Show vbModal
End If
If lstLoad_Forms.Text = "Fee Payment" Then
frmPay.Show vbModal
End If
If lstLoad_Forms.Text = "Print Situation Report" Then
frmReport.Show vbModal
End If
If lstLoad_Forms.Text = "About IBRUQ-TECH NIG" Then
frmAbout.Show vbModal
End If
If lstLoad_Forms.Text = "Exit" Then
If MsgBox("Exit Program?", vbYesNo + vbQuestion, "Exit") = vbYes Then
End
Else
Cancel = True
End If
End If
End With
End Sub

Private Sub mnuLogOut_Click()
lstLoad_Forms.Enabled = False
End Sub

Private Sub mnuLoin_Click()
frmLogin.Show vbModal
End Sub
