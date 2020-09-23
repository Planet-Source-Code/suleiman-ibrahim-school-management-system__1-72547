VERSION 5.00
Begin VB.Form frmPay_Summary 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   Icon            =   "frmPay_Summary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtStudent_No 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton cmdGet_Record 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   5175
   End
   Begin VB.OptionButton optAll 
      Caption         =   "SUMMARY OF ALL PAYMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   4935
   End
   Begin VB.OptionButton optIndi 
      Caption         =   "SUMMARY OF INDIVIDUAL PAYMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
End
Attribute VB_Name = "frmPay_Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGet_Record_Click()
Dim VarStudent_No As String
On Error Resume Next
VarStudent_No = txtStudent_No
If optIndi.Value = True Then
frmPay.DatIndividual.DatabaseName = App.Path & "\SMS.mdb"
frmPay.DatIndividual.RecordSource = "Select * from Fee where RegNo='" & VarStudent_No & "';"
frmPay.DatIndividual.Refresh
frmPay.Show
End If
If optAll.Value = True Then
frmPay.DatAllPay.DatabaseName = App.Path & "\SMS.mdb"
frmPay.DatAllPay.RecordSource = "Fee"
frmPay.DatAllPay.Refresh
frmPay.Show
'SendKeys "{Home}+{End}"
End If
frmPay_Summary.Hide
frmPay_Summary.Caption = "Checking Payment Details"
End Sub

Private Sub Form_Load()
frmPay_Summary.Caption = "Summary of Fee Payment"
End Sub

