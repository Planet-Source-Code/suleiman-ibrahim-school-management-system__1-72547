VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegist 
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   12045
   Icon            =   "frmRegist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00004080&
      Height          =   975
      Left            =   5640
      TabIndex        =   50
      Top             =   7800
      Width           =   6255
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search Record"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Search Record to Edit "
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdNew_Rec 
         Caption         =   "&New"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Enter New Record"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit_Rec 
         Caption         =   "&Edit"
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Search Record to Edit "
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete_Rec 
         Caption         =   "&Delete "
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
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Delete this Record?"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd_Rec 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "SaveNew Student Record"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate_Rec 
         Caption         =   "&Update"
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Save Edited Record"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdClearForm 
         Caption         =   "Clear Form"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Search Record to Edit "
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9240
      Top             =   9480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SMS_Proj\sms_database\SMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SMS_Proj\sms_database\SMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Registration"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   375
      Left            =   9960
      TabIndex        =   48
      Top             =   9000
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   9360
      TabIndex        =   47
      Top             =   9000
      Width           =   615
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   375
      Left            =   2040
      TabIndex        =   46
      Top             =   9000
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   375
      Left            =   1440
      TabIndex        =   45
      Top             =   9000
      Width           =   615
   End
   Begin VB.ComboBox cboProg_Applied 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      TabIndex        =   44
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox txtYear_Obt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   41
      Top             =   4800
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog cdl_Passport 
      Left            =   10080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Brawse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   35
      Top             =   2760
      Width           =   2655
   End
   Begin VB.ComboBox cboQualif 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      TabIndex        =   34
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox txtSch_Attend 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   33
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtSponsor_Address 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   7560
      Width           =   2655
   End
   Begin VB.TextBox txtSponsor_Name 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   31
      Top             =   7080
      Width           =   2655
   End
   Begin VB.TextBox txtMobile_No 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   29
      Top             =   6600
      Width           =   2655
   End
   Begin VB.TextBox txt_Address 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   5160
      Width           =   2655
   End
   Begin VB.ComboBox cbo_Occup 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   27
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txtLocal_Govt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   26
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox txtState_Origin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   25
      Top             =   3720
      Width           =   2655
   End
   Begin VB.ComboBox cboMarital_Status 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   24
      Top             =   3240
      Width           =   2295
   End
   Begin VB.ComboBox cbo_Religion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   23
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txt_Age 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   22
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox cbo_Sex 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   21
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtOther_Name 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   20
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtSur_Name 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   19
      Top             =   840
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Passport Size (120 X 140 pixels)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2415
      Left            =   7320
      TabIndex        =   39
      Top             =   240
      Width           =   2655
      Begin VB.PictureBox pic_Picture 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   480
         ScaleHeight     =   119
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   119
         TabIndex        =   40
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSMask.MaskEdBox mskAdm_Date 
      Height          =   375
      Left            =   7920
      TabIndex        =   42
      Top             =   5760
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSMask.MaskEdBox mskProg_Fee 
      Height          =   375
      Left            =   7920
      TabIndex        =   43
      Top             =   6480
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   """N""#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5115
      Top             =   4545
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegist.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegist.frx":0554
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegist.frx":0666
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegist.frx":0778
            Key             =   "Rectangle"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegist.frx":088A
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2640
      TabIndex        =   49
      Top             =   9000
      Width           =   6735
   End
   Begin VB.Label lblGenerate_No 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label4"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   38
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblPrefix_No 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   9480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label22 
      Caption         =   "Programme Fee:"
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
      TabIndex        =   36
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label21 
      Caption         =   "Admission Date:"
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
      TabIndex        =   30
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblAdm_No 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label19 
      Caption         =   "Programme Applied:"
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
      TabIndex        =   17
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "Year Obtained:"
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
      TabIndex        =   16
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label16 
      Caption         =   "Qualification Obtained:"
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
      TabIndex        =   15
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Last School Attended:"
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
      TabIndex        =   14
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Sponsor's Address:"
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
      Left            =   840
      TabIndex        =   13
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label Label13 
      Caption         =   "Sponsor's Name:"
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
      Left            =   840
      TabIndex        =   12
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Mobile Number:"
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
      Left            =   840
      TabIndex        =   11
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Contact Address:"
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
      Left            =   840
      TabIndex        =   10
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Occupation:"
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
      Left            =   840
      TabIndex        =   9
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Local Govt Area:"
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
      Left            =   840
      TabIndex        =   8
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "State of Origin:"
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
      Left            =   840
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Marital Status:"
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
      Left            =   840
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Religion:"
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
      Left            =   840
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Age:"
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
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Sex:"
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
      Left            =   840
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Other Names:"
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
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Surname:"
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
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Admission Number:"
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
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Program Menu"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuL2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuL3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Begin VB.Menu mnuPassport 
         Caption         =   "Passport"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "frmRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim SMS_DB As Database
    Dim rstRegist As Recordset
    Dim rstCheck_No As Recordset
    Dim VarPre_No As String
    Dim VarGenerate_No As String
    Dim VarCheck_No As String
    Dim strFileName As String
    Private Sub GetID()
    On Error GoTo ErrorTrap
    VarPre_No = "ICI/No/"
    lblPrefix_No = VarPre_No
    VarGenerate_No = VarCheck_No
    With rstCheck_No
    On Error Resume Next
    .MoveLast
    VarGenerate_No = Format(1, "000")
    lblGenerate_No = VarGenerate_No
    lblGenerate_No = Format(CDbl(![Vacant_No]) + 1, "000")
    lblAdm_No = lblPrefix_No & lblGenerate_No
    End With
    Exit Sub
ErrorTrap:
     MsgBox Err.Description, vbInformation, "Error"
    End Sub
    Private Sub ClearForm()
    On Error Resume Next
     lblAdm_No = ""
     txtSur_Name = ""
     txtOther_Name = ""
     cbo_Sex = ""
     txt_Age = ""
     cbo_Religion = ""
     cboMarital_Status = ""
     txtState_Origin = ""
     txtLocal_Govt = ""
     cbo_Occup = ""
     txt_Address = ""
     txtMobile_No = ""
     pic_Picture.Picture = LoadPicture(App.Path & "\Passport\NoPic.jpg")
     txtSponsor_Name = ""
     txtSponsor_Address = ""
     txtSch_Attend = ""
     cboQualif = ""
     txtYear_Obt = ""
     strFileName = ""
     cboProg_Applied = ""
     mskProg_Fee = ""
     mskAdm_Date = ""
    End Sub
    Private Sub cmdAdd_Rec_Click()
    On Error Resume Next
    With rstRegist
    '.MoveNext
    .AddNew
    !Adm_No = lblAdm_No
    !SurName = txtSur_Name
    ![Other Name] = txtOther_Name
    !Sex = cbo_Sex
    !Age = txt_Age
    !Religion = cbo_Religion
    ![Marital Status] = cboMarital_Status
    !State = txtState_Origin
    !LGA = txtLocal_Govt
    !Occupation = cbo_Occup
    ![Contact Address] = txt_Address
    ![Mobile No] = txtMobile_No
    !Sponsor = txtSponsor_Name
    ![Sponsor Address] = txtSponsor_Address
    ![Last School Attended] = txtSch_Attend
    !Qualification = cboQualif
    ![Year Obtained] = txtYear_Obt
    !Passport = strFileName
    !Program = cboProg_Applied
    ![Programme Fee] = mskProg_Fee
    !Adm_Date = mskAdm_Date
    .Update
    .Bookmark = .LastModified
    End With
    With rstCheck_No
    On Error Resume Next
    '.MoveNext
    .AddNew
    !Vacant_No = lblGenerate_No
    .Update
    .Bookmark = .LastModified
    End With
    ClearForm
    End Sub
    
    Private Sub cmdBrowse_Click()
    cdl_Passport.DialogTitle = "Upload Passport"
    cdl_Passport.InitDir = App.Path
    cdl_Passport.Filter = "All Files (*.*) |*.*|JPEG Files (*.JPG) |*.JPG|"
    cdl_Passport.ShowOpen
    strFileName = cdl_Passport.FileName
    pic_Picture.Picture = LoadPicture(strFileName)
    End Sub
    
Private Sub cmdClearForm_Click()
cmdSearch.Visible = True
cmdClearForm.Visible = False
Call ClearForm
End Sub

    Private Sub cmdDelete_Rec_Click()
    If MsgBox("Delete this Record?", vbYesNo + vbQuestion, "delete") = vbNo Then
    Exit Sub
    End If
    With rstRegist
    .Delete
    Call ClearForm
    End With
    End Sub

    Private Sub cmdEdit_Rec_Click()
    On Error Resume Next
    cmdUpdate_Rec.Visible = True
    cmdEdit_Rec.Visible = False
    Dim StrSearch As String
    StrSearch = InputBox("Adminssion Number:", "Find Student")
    On Error Resume Next
    With rstRegist
    .Index = "Adm_No"
    .Seek "=", StrSearch
    If .NoMatch Then
    MsgBox "No Record", vbInformation, "Find Student"
    Exit Sub
    Else
    Get_Record
    End If
    End With
    txtFullName.SetFocus
    End Sub
    Private Sub Get_Record()
    On Error Resume Next
    With rstRegist
    lblAdm_No = !Adm_No
    txtSur_Name = !SurName
     txtOther_Name = ![Other Name]
     cbo_Sex = !Sex
     txt_Age = !Age
     cbo_Religion = !Religion
     cboMarital_Status = ![Marital Status]
     txtState_Origin = !State
     txtLocal_Govt = !LGA
     cbo_Occup = !Occupation
     txt_Address = ![Contact Address]
     txtMobile_No = ![Mobile No]
     txtSponsor_Name = !Sponsor
     txtSponsor_Address = ![Sponsor Address]
     txtSch_Attend = ![Last School Attended]
     cboQualif = !Qualification
     pic_Picture.Picture = LoadPicture(!Passport)
     txtYear_Obt = ![Year Obtained]
     cboProg_Applied = !Program
     mskProg_Fee = ![Programme Fee]
     mskAdm_Date = ![Adm_Date]
     End With
    End Sub

 Private Sub cmdFirst_Click()
 On Error Resume Next
    With rstRegist
    .MoveFirst
    Call Get_Record
    End With
    End Sub
    Private Sub cmdLast_Click()
    On Error Resume Next
    With rstRegist
    .MoveLast
    Call Get_Record
    End With
    End Sub

    Private Sub cmdNew_Rec_Click()
    On Error Resume Next
    cmdNew_Rec.Visible = False
    cmdAdd_Rec.Visible = True
    ClearForm
    GetID
    End Sub

    Private Sub UpdateRec()
    With rstRegist
     !Adm_No = lblAdm_No
     !SurName = txtSur_Name
     ![Other Name] = txtOther_Name
      !Sex = cbo_Sex
      !Age = txt_Age
      !Religion = cbo_Religion
      ![Marital Status] = cboMarital_Status
      !State = txtState_Origin
      !LGA = txtLocal_Govt
      !Occupation = cbo_Occup
      ![Contact Address] = txt_Address
      ![Mobile No] = txtMobile_No
      !Sponsor = txtSponsor_Name
      ![Sponsor Address] = txtSponsor_Address
      ![Last School Attended] = txtSch_Attend
      !Qualification = cboQualif
      ![Year Obtained] = txtYear_Obt
      !Passport = strFileName
      !Program = cboProg_Applied
      ![Programme Fee] = mskProg_Fee
      ![Adm_Date] = mskAdm_Date
     End With
        End Sub

 Private Sub cmdNext_Click()
 On Error Resume Next
    With rstRegist
    .MoveNext
    If .EOF Then
    .MoveLast
    End If
    Call Get_Record
    End With
    End Sub
    Private Sub cmdPrevious_Click()
    On Error Resume Next
    With rstRegist
    .MovePrevious
    If .BOF Then
    .MoveFirst
    End If
   Call Get_Record
    End With
    End Sub

Private Sub cmdRegister_Report_Click()
'On Error Resume Next

End Sub

Private Sub cmdSearch_Click()
cmdSearch.Visible = False
cmdClearForm.Visible = True
Call cmdEdit_Rec_Click
End Sub

    Private Sub cmdUpdate_Rec_Click()
    cmdEdit_Rec.Visible = True
    cmdUpdate_Rec.Visible = False
    With rstRegist
    '.MoveNext
    .Edit
    Call UpdateRec
    .Update
    .Bookmark = .LastModified
    End With
    Call ClearForm
    End Sub

    Private Sub Form_Load()
    Set SMS_DB = OpenDatabase(App.Path & "\SMS.mdb", False, False)
    Set rstRegist = SMS_DB.OpenRecordset("Registration")
    Set rstCheck_No = SMS_DB.OpenRecordset("Check_No")
    Call AddMode
    frmRegist.Caption = "Registering Student"
    End Sub
    
    
    Private Sub AddMode()
    cbo_Occup.AddItem "Civil Servant"
    cbo_Occup.AddItem "Business Man"
    cbo_Occup.AddItem "Student"
    cbo_Occup.AddItem "Others"
    
    cbo_Religion.AddItem "Christianity"
    cbo_Religion.AddItem "Islam"
    cbo_Religion.AddItem "Others"
    
    cbo_Sex.AddItem "Female"
    cbo_Sex.AddItem "Male"
     cboMarital_Status.AddItem "Single"
      cboMarital_Status.AddItem "Married"
       cboMarital_Status.AddItem "Widow"
       
    cboProg_Applied.AddItem "Certificate"
        cboProg_Applied.AddItem "Diploma"
            cboProg_Applied.AddItem "Special"
    
     cboQualif.AddItem "SSCE"
     cboQualif.AddItem "NECO"
     cboQualif.AddItem "ND/OND/NCE"
     cboQualif.AddItem "Degree/HND"
     cboQualif.AddItem "Others"
    End Sub
    
Private Sub mnuClose_Click()
Me.Hide
End Sub

Private Sub mnuDelete_Click()
cmdDelete_Rec_Click
End Sub

Private Sub mnuNew_Click()
cmdNew_Rec_Click
End Sub

Private Sub mnuOpen_Click()
cmdEdit_Rec_Click
End Sub

Private Sub mnuPassport_Click()
cmdBrowse_Click
End Sub

'Private Sub mnuPrint_Click()
'frmRegist.PrintForm
'End Sub

Private Sub mnuSave_Click()
cmdAdd_Rec_Click
End Sub

Private Sub mnuUpdate_Click()
cmdUpdate_Rec_Click
End Sub
