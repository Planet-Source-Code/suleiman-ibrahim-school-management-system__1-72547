VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   9180
      Left            =   0
      Picture         =   "frmAbout.frx":0442
      Top             =   0
      Width           =   10050
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmAbout.Caption = "About IBRUQ-TECH NIG"
End Sub
