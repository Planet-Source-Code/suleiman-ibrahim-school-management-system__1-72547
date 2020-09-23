VERSION 5.00
Begin VB.Form frmReport 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPay_By_Trans 
      Appearance      =   0  'Flat
      Caption         =   "Payment By Transaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   4080
      Width           =   7455
   End
   Begin VB.CommandButton cmdIndRpt 
      Appearance      =   0  'Flat
      Caption         =   "Print Individual Fee Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   7455
   End
   Begin VB.CommandButton cmdDiplomaRpt 
      Appearance      =   0  'Flat
      Caption         =   "Print Diploma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   7455
   End
   Begin VB.CommandButton cmdEngRpt 
      Appearance      =   0  'Flat
      Caption         =   "Print Engineering"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   7455
   End
   Begin VB.CommandButton cmdCertRpt 
      Appearance      =   0  'Flat
      Caption         =   "Print Certificate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   7455
   End
   Begin VB.CommandButton cmdRegister_Report 
      Appearance      =   0  'Flat
      Caption         =   "Print Registration "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   7455
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCertRpt_Click()
On Error Resume Next
Dim Slip As String
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Set db = New ADODB.Connection
db.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data source=" & App.Path & "\SMS.mdb"
Slip = InputBox("Enter Registration No. To Print")
Set rs = New ADODB.Recordset
rs.Open "Select * From Certificate Where [Reg No] =" & "'" & Slip & "'", db, adOpenDynamic, adLockOptimistic

Set rptCertificate.DataSource = rs
rptCertificate.Refresh
rptCertificate.Show vbModal
End Sub

Private Sub cmdDiplomaRpt_Click()
On Error Resume Next
Dim Slip As String
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Set db = New ADODB.Connection
db.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data source=" & App.Path & "\SMS.mdb"
Slip = InputBox("Enter Registration No. To Print")
Set rs = New ADODB.Recordset
rs.Open "Select * From Diploma Where [Reg No] =" & "'" & Slip & "'", db, adOpenDynamic, adLockOptimistic

Set rptDiploma.DataSource = rs
rptDiploma.Refresh
rptDiploma.Show vbModal

End Sub

Private Sub cmdEngRpt_Click()
On Error Resume Next
Dim Slip As String
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Set db = New ADODB.Connection
db.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data source=" & App.Path & "\SMS.mdb"
Slip = InputBox("Enter Registration No. To Print")
Set rs = New ADODB.Recordset
rs.Open "Select * From Engineering Where [Reg No] =" & "'" & Slip & "'", db, adOpenDynamic, adLockOptimistic

Set rptEngineering.DataSource = rs
rptEngineering.Refresh
rptEngineering.Show vbModal
End Sub

Private Sub cmdIndRpt_Click()
On Error Resume Next
Dim Slip As String
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Set db = New ADODB.Connection
db.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data source=" & App.Path & "\SMS.mdb"
Slip = InputBox("Enter Registration No. To Print")
Set rs = New ADODB.Recordset
rs.Open "Select * From Fee Where [RegNo] =" & "'" & Slip & "'", db, adOpenDynamic, adLockOptimistic

Set rptPayment.DataSource = rs
rptPayment.Refresh
rptPayment.Show vbModal
End Sub

Private Sub cmdPay_By_Trans_Click()
On Error Resume Next
Dim Slip As String
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Set db = New ADODB.Connection
db.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data source=" & App.Path & "\SMS.mdb"
Slip = InputBox("Enter Transaction No. To Print")
Set rs = New ADODB.Recordset
rs.Open "Select * From Fee Where [Trans_No] =" & "'" & Slip & "'", db, adOpenDynamic, adLockOptimistic

Set rptPayment.DataSource = rs
rptPayment.Refresh
rptPayment.Show vbModal
End Sub

Private Sub cmdRegister_Report_Click()
On Error Resume Next
Dim Slip As String
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Set db = New ADODB.Connection
db.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data source=" & App.Path & "\SMS.mdb"
Slip = InputBox("Enter Admission No. To Print")
Set rs = New ADODB.Recordset
rs.Open "Select * From Registration Where [Adm_No] =" & "'" & Slip & "'", db, adOpenDynamic, adLockOptimistic

Set DataReport_Register.DataSource = rs
DataReport_Register.Refresh
DataReport_Register.Show vbModal
End Sub


Private Sub Form_Load()
frmReport.Caption = "Print Report"
End Sub
