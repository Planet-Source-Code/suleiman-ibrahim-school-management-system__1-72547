VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2910
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4845
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1719.324
   ScaleMode       =   0  'User
   ScaleWidth      =   4549.193
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   630
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C000&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2685
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0C000&
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1320
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
        If txtPassword = "school" Then
        LoginSucceeded = True
        Me.Hide
        frmMain.lstLoad_Forms.Enabled = True
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Load()
frmLogin.Caption = "Enter Your Password"
End Sub
