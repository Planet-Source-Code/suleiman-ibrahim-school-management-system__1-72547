VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPay 
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14295
   Icon            =   "frmPay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   14295
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "FEE PAYMENT"
      TabPicture(0)   =   "frmPay.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTrans_No"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAmount_Pay"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdNewPayt"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSaveNew"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDelete"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdAddPayment"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSaveUpdate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "INDIVIDUAL PAYMENT REPORT"
      TabPicture(1)   =   "frmPay.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "ALL PAYMENT REPORT"
      TabPicture(2)   =   "frmPay.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command2"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdSaveUpdate 
         Caption         =   "Post Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   38
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddPayment 
         Caption         =   "Update Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   37
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Record"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   36
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "Post New Paid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   35
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdNewPayt 
         Caption         =   "New Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   34
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Check Last Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -63480
         TabIndex        =   33
         Top             =   4800
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Check Last Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -63600
         TabIndex        =   32
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   -74760
         TabIndex        =   24
         Top             =   480
         Width           =   13455
         Begin VB.Data DatAllPay 
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   240
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3840
            Visible         =   0   'False
            Width           =   12135
         End
         Begin MSFlexGridLib.MSFlexGrid msfAll_Pay 
            Bindings        =   "frmPay.frx":0496
            Height          =   3495
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   13
            FixedCols       =   0
            GridLines       =   2
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   -74760
         TabIndex        =   23
         Top             =   480
         Width           =   13455
         Begin VB.Data DatIndividual 
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   240
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3840
            Visible         =   0   'False
            Width           =   12135
         End
         Begin MSFlexGridLib.MSFlexGrid msfIndividual 
            Bindings        =   "frmPay.frx":04AE
            Height          =   3495
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   13
            FixedCols       =   0
            GridLines       =   2
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4215
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   10455
         Begin MSMask.MaskEdBox mskAmount_Pay 
            Height          =   375
            Left            =   5880
            TabIndex        =   22
            Top             =   3360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskTrans_Date 
            Height          =   375
            Left            =   2160
            TabIndex        =   20
            Top             =   3360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd-mmm-yy"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            Caption         =   "MM/DD/YY"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   31
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label lblProgram 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   29
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Program:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   28
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Amount To Pay:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   21
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label lblOutst_Payt 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   19
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label lblTotal_Paid 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   18
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label lblLast_Payt 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   17
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label lblProg_Fee 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   16
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblLast_Pay_No 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   15
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label lblPay_No 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   14
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblOther_Name 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7800
            TabIndex        =   13
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblSur_Name 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5760
            TabIndex        =   12
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblReg_No 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Last Payment Number:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   10
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label9 
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   9
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Reg Number:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Student Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   7
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Payment Number:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Program Fee:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   5
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Last Payment:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   4
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Total Payment:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Out Standing:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   2
            Top             =   2760
            Width           =   1215
         End
      End
      Begin VB.Label lblAmount_Pay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4440
         TabIndex        =   30
         Top             =   4800
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label lblTrans_No 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         TabIndex        =   27
         Top             =   4800
         Visible         =   0   'False
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SMS_DB As Database
 Dim rstRegist As Recordset
 Dim rstFee As Recordset


Private Sub cmdAddPayment_Click()
On Error Resume Next
Dim StrSearch As String
Dim StrStudentNo As String
Dim StrPaymentNo As String
StrStudentNo = InputBox("Enter Student Number:", "Search Record")
StrPaymentNo = InputBox("Enter Payment Number:", "Search Record")
StrSearch = StrStudentNo & StrPaymentNo
With rstFee
.Index = "Trans_No"
.Seek "=", StrSearch
If .NoMatch Then
MsgBox "Record Not Found", vbInformation, "Search"
End If
End With
LoadData
Dim D As Currency, C As Currency
D = CCur(lblLast_Payt)
C = D
lblAmount_Pay = CCur(C)
    With rstFee
    lblLast_Pay_No = !Pay_No
    lblPay_No = (!Pay_No) + 1
    End With
End Sub
Private Sub LoadData()
On Error Resume Next
With rstFee
 lblTrans_No = !Trans_No
 mskTrans_Date = ""
 lblReg_No = !RegNo
 lblSur_Name = ![Sur Name]
 lblOther_Name = ![Other Name]
 lblProgram = !Programme
 lblProg_Fee = ![Fee Charged]
 mskAmount_Pay = ![Current Pay]
 lblLast_Payt = ![Last Pay]
 lblTotal_Paid = ![Total Paid]
 lblOutst_Payt = ![Out StandPay]
 lblLast_Pay_No = !Pay_No
 lblLast_Pay_No = !Last_Pay_No
 End With
End Sub
Private Sub cmdNewPayt_Click()
Dim VarPay_No As String
Dim StrSearch As String
    StrSearch = InputBox("Adminssion Number:", "Find Student")
    On Error Resume Next
    With rstRegist
    .Index = "Adm_No"
    .Seek "=", StrSearch
    If .NoMatch Then
    MsgBox "There is No Record", vbInformation, "Find Student"
    Exit Sub
    Else
   lblReg_No = !Adm_No
   lblSur_Name = !SurName
   lblOther_Name = ![Other Name]
   lblProg_Fee = ![Programme Fee]
   lblProgram = !Program
    End If
    End With
    With rstFee
    VarPay_No = Format(1, 0)
    lblPay_No = VarPay_No
    End With
    
End Sub

Private Sub cmdSaveNew_Click()
Call ComputeAmount
On Error GoTo ErrTrap
With rstFee
.AddNew
!Trans_No = lblReg_No & lblPay_No
!Payment_Date = mskTrans_Date
!RegNo = lblReg_No
![Sur Name] = lblSur_Name
![Other Name] = lblOther_Name
!Programme = lblProgram
![Fee Charged] = lblProg_Fee
![Current Pay] = mskAmount_Pay
![Last Pay] = lblLast_Payt
![Total Paid] = lblTotal_Paid
![Out StandPay] = lblOutst_Payt
!Pay_No = lblPay_No
!Last_Pay_No = lblPay_No
.Update
.Bookmark = .LastModified
End With
Exit Sub
ErrTrap:
MsgBox Err.Description, vbInformation, "Error"
End Sub
Private Sub ComputeAmount()
Dim A As Currency, B As Currency, C As Currency
Dim D As Currency, E As Currency, F As Currency
Dim Counter As Integer
On Error Resume Next
A = CCur(mskAmount_Pay)
B = A
lblAmount_Pay = CCur(B)
If lblAmount_Pay = Visible Then
mskAmount_Pay = 0
C = B
lblLast_Payt = CCur(C)
D = lblTotal_Paid
lblTotal_Paid = D + lblLast_Payt
F = CCur(lblProg_Fee)
E = F - Val(lblTotal_Paid)
lblOutst_Payt = Format(CCur(E), " #,####0.00")
End If
End Sub

Private Sub cmdSaveUpdate_Click()
Call ComputeAmount
On Error GoTo ErrTrap
With rstFee
.AddNew
!Trans_No = lblReg_No & lblPay_No
!Payment_Date = mskTrans_Date
!RegNo = lblReg_No
![Sur Name] = lblSur_Name
![Other Name] = lblOther_Name
!Programme = lblProgram
![Fee Charged] = lblProg_Fee
![Current Pay] = mskAmount_Pay
![Last Pay] = lblLast_Payt
![Total Paid] = lblTotal_Paid
![Out StandPay] = lblOutst_Payt
!Pay_No = lblPay_No
!Last_Pay_No = lblPay_No
.Update
.Bookmark = .LastModified
End With
Exit Sub
ErrTrap:
MsgBox Err.Description, vbInformation, "Error"
End Sub

Private Sub Command1_Click()
frmPay_Summary.optAll.Visible = False
frmPay_Summary.optIndi = True
frmPay_Summary.txtStudent_No.Visible = True
frmPay_Summary.txtStudent_No = ""
frmPay_Summary.Show vbModal
End Sub

Private Sub Command2_Click()
frmPay_Summary.optIndi.Visible = False
frmPay_Summary.txtStudent_No.Visible = False
frmPay_Summary.optAll.Visible = True
frmPay_Summary.Show vbModal
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
Set SMS_DB = OpenDatabase(App.Path & "\SMS.mdb", False, False)
Set rstRegist = SMS_DB.OpenRecordset("Registration")
Set rstFee = SMS_DB.OpenRecordset("Fee")
frmPay.Caption = "Making Payment"
Call GIndiv
Call GAll
End Sub
Private Sub GIndiv()
msfIndividual.ColWidth(0) = 1300
msfIndividual.ColWidth(1) = 1200
msfIndividual.ColWidth(2) = 1300
msfIndividual.ColWidth(3) = 1500
msfIndividual.ColWidth(4) = 1500
msfIndividual.ColWidth(5) = 1100
msfIndividual.ColWidth(6) = 1200
msfIndividual.ColWidth(7) = 0
msfIndividual.ColWidth(8) = 1200
msfIndividual.ColWidth(9) = 1300
msfIndividual.ColWidth(10) = 1200
msfIndividual.ColWidth(11) = 0
msfIndividual.ColWidth(12) = 0
End Sub
Private Sub GAll()
msfAll_Pay.ColWidth(0) = 1300
msfAll_Pay.ColWidth(1) = 1200
msfAll_Pay.ColWidth(2) = 1300
msfAll_Pay.ColWidth(3) = 1500
msfAll_Pay.ColWidth(4) = 1500
msfAll_Pay.ColWidth(5) = 1100
msfAll_Pay.ColWidth(6) = 1200
msfAll_Pay.ColWidth(7) = 0
msfAll_Pay.ColWidth(8) = 1200
msfAll_Pay.ColWidth(9) = 1300
msfAll_Pay.ColWidth(10) = 1200
msfAll_Pay.ColWidth(11) = 0
msfAll_Pay.ColWidth(12) = 0
End Sub
