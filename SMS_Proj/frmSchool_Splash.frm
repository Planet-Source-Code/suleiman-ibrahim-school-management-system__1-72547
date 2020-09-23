VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSchool_Splash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4680
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSchool_Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4395
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   7080
      Begin VB.Timer loadTimer 
         Interval        =   1000
         Left            =   6480
         Top             =   3360
      End
      Begin MSComctlLib.ProgressBar SMS_ProgressBar 
         Height          =   255
         Left            =   3000
         TabIndex        =   0
         Top             =   3000
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Timer progressTimer 
         Interval        =   1000
         Left            =   6360
         Top             =   2400
      End
      Begin VB.Image Image1 
         Height          =   1515
         Left            =   120
         Picture         =   "frmSchool_Splash.frx":0442
         Top             =   120
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Wait for Programme to Load"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   6
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   " Warning : This Product is Licensed to IBRUQ-TECH NIG  Use of pirated copy is illegal."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   3960
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "Copyright (C) IBRUQ: Product is copyrighted in the year 2008"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   3600
         Width           =   5175
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   3
         Top             =   1800
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmSchool_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer


Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = "IBRUQ for School Management System"
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub




Private Sub progressTimer_Timer()
    i = Rnd() * 30
    If SMS_ProgressBar.Value < 80 Then
        If SMS_ProgressBar.Value + i < 80 Then
            SMS_ProgressBar.Value = SMS_ProgressBar.Value + i
        Else
            SMS_ProgressBar.Value = 80
        End If
            Else
    Unload Me
    frmMain.Show
      End If
      End Sub

