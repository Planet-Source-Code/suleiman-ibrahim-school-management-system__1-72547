VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmResult 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   13996
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CERTIFICATE"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(3)=   "Command2"
      Tab(0).Control(4)=   "Command4"
      Tab(0).Control(5)=   "Command5"
      Tab(0).Control(6)=   "Command6"
      Tab(0).Control(7)=   "Command17"
      Tab(0).Control(8)=   "Command18"
      Tab(0).Control(9)=   "Command19"
      Tab(0).Control(10)=   "Command20"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "DIPLOMA"
      TabPicture(1)   =   "frmResult.frx":0442
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label27"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Command7"
      Tab(1).Control(3)=   "Command8"
      Tab(1).Control(4)=   "Command9"
      Tab(1).Control(5)=   "Command10"
      Tab(1).Control(6)=   "Command11"
      Tab(1).Control(7)=   "Command21"
      Tab(1).Control(8)=   "Command22"
      Tab(1).Control(9)=   "Command23"
      Tab(1).Control(10)=   "Command27"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "ENGINEERING"
      TabPicture(2)   =   "frmResult.frx":045E
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label28"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command12"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command13"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command14"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command15"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Command16"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Command24"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Command25"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Command26"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Command28"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.CommandButton Command28 
         Caption         =   "<"
         Height          =   375
         Left            =   1440
         TabIndex        =   129
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command27 
         Caption         =   "<"
         Height          =   375
         Left            =   -73560
         TabIndex        =   128
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command26 
         Caption         =   "|<"
         Height          =   375
         Left            =   960
         TabIndex        =   126
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command25 
         Caption         =   ">"
         Height          =   375
         Left            =   7080
         TabIndex        =   125
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command24 
         Caption         =   ">|"
         Height          =   375
         Left            =   7560
         TabIndex        =   124
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command23 
         Caption         =   "|<"
         Height          =   375
         Left            =   -74040
         TabIndex        =   122
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command22 
         Caption         =   ">"
         Height          =   375
         Left            =   -67920
         TabIndex        =   121
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command21 
         Caption         =   ">|"
         Height          =   375
         Left            =   -67440
         TabIndex        =   120
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command20 
         Caption         =   ">|"
         Height          =   375
         Left            =   -67320
         TabIndex        =   119
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command19 
         Caption         =   ">"
         Height          =   375
         Left            =   -67800
         TabIndex        =   118
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command18 
         Caption         =   "|<"
         Height          =   375
         Left            =   -73920
         TabIndex        =   117
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command17 
         Caption         =   "<"
         Height          =   375
         Left            =   -73440
         TabIndex        =   116
         Top             =   7320
         Width           =   495
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Delete"
         Height          =   495
         Left            =   6720
         TabIndex        =   111
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Edit Result"
         Height          =   495
         Left            =   5280
         TabIndex        =   110
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Search Result"
         Height          =   495
         Left            =   3840
         TabIndex        =   109
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Compute Result"
         Height          =   495
         Left            =   2400
         TabIndex        =   108
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         Caption         =   "New Result"
         Height          =   495
         Left            =   960
         TabIndex        =   107
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -68280
         TabIndex        =   106
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Edit Result"
         Height          =   495
         Left            =   -69720
         TabIndex        =   105
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Search Result"
         Height          =   495
         Left            =   -71160
         TabIndex        =   104
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Compute Result"
         Height          =   495
         Left            =   -72600
         TabIndex        =   103
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "New Result"
         Height          =   495
         Left            =   -74040
         TabIndex        =   102
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -68160
         TabIndex        =   95
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Edit Result"
         Height          =   495
         Left            =   -69600
         TabIndex        =   94
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Search Result"
         Height          =   495
         Left            =   -71040
         TabIndex        =   93
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Compute Result"
         Height          =   495
         Left            =   -72480
         TabIndex        =   91
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New Result"
         Height          =   495
         Left            =   -73920
         TabIndex        =   90
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Height          =   6255
         Left            =   960
         TabIndex        =   55
         Top             =   600
         Width           =   7095
         Begin VB.TextBox txtEProj 
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
            Left            =   5640
            TabIndex        =   89
            Top             =   4200
            Width           =   855
         End
         Begin VB.TextBox txtESystUp 
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
            Left            =   5640
            TabIndex        =   88
            Top             =   3720
            Width           =   855
         End
         Begin VB.TextBox txtEMaint 
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
            Left            =   5640
            TabIndex        =   87
            Top             =   3240
            Width           =   855
         End
         Begin VB.TextBox txtETroubSh 
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
            Left            =   5640
            TabIndex        =   86
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox txtESoftIN 
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
            Left            =   5640
            TabIndex        =   85
            Top             =   2280
            Width           =   855
         End
         Begin VB.TextBox txtEHardIN 
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
            TabIndex        =   84
            Top             =   4200
            Width           =   855
         End
         Begin VB.TextBox txtEHardIDN 
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
            TabIndex        =   83
            Top             =   3720
            Width           =   855
         End
         Begin VB.TextBox txtEDos 
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
            TabIndex        =   82
            Top             =   3240
            Width           =   855
         End
         Begin VB.TextBox txtEWin 
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
            TabIndex        =   81
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox txtEAppre 
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
            TabIndex        =   80
            Top             =   2280
            Width           =   855
         End
         Begin MSMask.MaskEdBox mskEComp_Date 
            Height          =   255
            Left            =   2640
            TabIndex        =   79
            Top             =   1800
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
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
         Begin VB.Label lblEOtherName 
            BackColor       =   &H8000000E&
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
            Left            =   4440
            TabIndex        =   114
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblETotal 
            BackColor       =   &H8000000E&
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
            Left            =   3120
            TabIndex        =   78
            Top             =   4800
            Width           =   2415
         End
         Begin VB.Label lblEAverage 
            BackColor       =   &H8000000E&
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
            Left            =   3120
            TabIndex        =   77
            Top             =   5280
            Width           =   2415
         End
         Begin VB.Label lblEClass 
            BackColor       =   &H8000000E&
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
            Left            =   3120
            TabIndex        =   76
            Top             =   5760
            Width           =   2415
         End
         Begin VB.Label Label47 
            Caption         =   "Classification:"
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
            Left            =   1680
            TabIndex        =   75
            Top             =   5760
            Width           =   1455
         End
         Begin VB.Label Label46 
            Caption         =   "Average Score:"
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
            Left            =   1680
            TabIndex        =   74
            Top             =   5280
            Width           =   1455
         End
         Begin VB.Label Label45 
            Caption         =   "Total Score:"
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
            Left            =   1680
            TabIndex        =   73
            Top             =   4800
            Width           =   1215
         End
         Begin VB.Label Label44 
            Caption         =   "Project:"
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
            Left            =   3840
            TabIndex        =   72
            Top             =   4200
            Width           =   1695
         End
         Begin VB.Label Label43 
            Caption         =   "System Upgrading:"
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
            Left            =   3840
            TabIndex        =   71
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Label Label42 
            Caption         =   "Maintenance:"
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
            Left            =   3840
            TabIndex        =   70
            Top             =   3240
            Width           =   1695
         End
         Begin VB.Label Label41 
            Caption         =   "Trouble Shooting:"
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
            Left            =   3840
            TabIndex        =   69
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label Label40 
            Caption         =   "Hardware Installation:"
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
            Left            =   600
            TabIndex        =   68
            Top             =   4200
            Width           =   1935
         End
         Begin VB.Label Label39 
            Caption         =   "Software Installation:"
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
            Left            =   3840
            TabIndex        =   67
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label38 
            Caption         =   "Hardware Identification:"
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
            Left            =   600
            TabIndex        =   66
            Top             =   3720
            Width           =   2055
         End
         Begin VB.Label Label37 
            Caption         =   "MS DOS:"
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
            Left            =   600
            TabIndex        =   65
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label Label36 
            Caption         =   "MS Windows:"
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
            Left            =   600
            TabIndex        =   64
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label Label35 
            Caption         =   "Appreciation:"
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
            Left            =   600
            TabIndex        =   63
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label lblESname 
            BackColor       =   &H8000000E&
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
            Left            =   2640
            TabIndex        =   62
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblEAdm 
            BackColor       =   &H8000000E&
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
            Left            =   2640
            TabIndex        =   61
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label lblEProgram 
            BackColor       =   &H8000000E&
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
            Left            =   2640
            TabIndex        =   60
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label32 
            Caption         =   "Completion Date:"
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
            Left            =   960
            TabIndex        =   59
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label31 
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
            Left            =   960
            TabIndex        =   58
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label30 
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
            Left            =   960
            TabIndex        =   57
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label29 
            Caption         =   "Programme:"
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
            Left            =   960
            TabIndex        =   56
            Top             =   1320
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6255
         Left            =   -74040
         TabIndex        =   26
         Top             =   600
         Width           =   7095
         Begin VB.TextBox txtDSeminar 
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
            Left            =   5280
            TabIndex        =   54
            Top             =   4320
            Width           =   735
         End
         Begin VB.TextBox txtDVB 
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
            Left            =   5280
            TabIndex        =   53
            Top             =   3840
            Width           =   735
         End
         Begin VB.TextBox txtDCorelDraw 
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
            Left            =   5280
            TabIndex        =   52
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox txtDDBMS 
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
            Left            =   5280
            TabIndex        =   51
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txtDAccess 
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
            Left            =   5280
            TabIndex        =   50
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox txtDPPt 
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
            Left            =   2760
            TabIndex        =   49
            Top             =   4320
            Width           =   855
         End
         Begin VB.TextBox txtDExcel 
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
            Left            =   2760
            TabIndex        =   48
            Top             =   3840
            Width           =   855
         End
         Begin VB.TextBox txtDWord 
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
            Left            =   2760
            TabIndex        =   47
            Top             =   3360
            Width           =   855
         End
         Begin VB.TextBox txtDWin 
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
            Left            =   2760
            TabIndex        =   46
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox txtDAppre 
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
            Left            =   2760
            TabIndex        =   45
            Top             =   2400
            Width           =   855
         End
         Begin MSMask.MaskEdBox mskDComp_Date 
            Height          =   255
            Left            =   3000
            TabIndex        =   44
            Top             =   1800
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
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
         Begin VB.Label lblDOther 
            BackColor       =   &H8000000E&
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
            Left            =   4560
            TabIndex        =   113
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label56 
            Caption         =   "Classification:"
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
            Left            =   1920
            TabIndex        =   101
            Top             =   5640
            Width           =   1575
         End
         Begin VB.Label Label55 
            Caption         =   "Average Score:"
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
            Left            =   1920
            TabIndex        =   100
            Top             =   5280
            Width           =   1575
         End
         Begin VB.Label Label54 
            Caption         =   "Total Score:"
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
            Left            =   1920
            TabIndex        =   99
            Top             =   4920
            Width           =   1575
         End
         Begin VB.Label lblDClass 
            BackColor       =   &H80000009&
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
            Left            =   3480
            TabIndex        =   98
            Top             =   5640
            Width           =   2055
         End
         Begin VB.Label lblDAverage 
            BackColor       =   &H80000009&
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
            Left            =   3480
            TabIndex        =   97
            Top             =   5280
            Width           =   2055
         End
         Begin VB.Label lblDTotal 
            BackColor       =   &H80000009&
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
            Left            =   3480
            TabIndex        =   96
            Top             =   4920
            Width           =   2055
         End
         Begin VB.Label lblDANum 
            BackColor       =   &H8000000E&
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
            Left            =   3000
            TabIndex        =   43
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label lblDSName 
            BackColor       =   &H8000000E&
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
            Left            =   3000
            TabIndex        =   42
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblDProg 
            BackColor       =   &H8000000E&
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
            Left            =   3000
            TabIndex        =   41
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label25 
            Caption         =   "MS Access:"
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
            TabIndex        =   40
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label24 
            Caption         =   "DBMS:"
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
            TabIndex        =   39
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label23 
            Caption         =   "CorelDraw:"
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
            TabIndex        =   38
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Intro To VB:"
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
            TabIndex        =   37
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Seminar:"
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
            TabIndex        =   36
            Top             =   4320
            Width           =   1335
         End
         Begin VB.Label Label19 
            Caption         =   "Appreciation:"
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
            Left            =   1200
            TabIndex        =   35
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Label18 
            Caption         =   "MS Windows:"
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
            Left            =   1200
            TabIndex        =   34
            Top             =   2880
            Width           =   1695
         End
         Begin VB.Label Label16 
            Caption         =   "Word Processing:"
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
            Left            =   1200
            TabIndex        =   33
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label Label15 
            Caption         =   "MS Excel:"
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
            Left            =   1200
            TabIndex        =   32
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "MS PowerPoint:"
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
            Left            =   1200
            TabIndex        =   31
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Programme:"
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
            Left            =   1320
            TabIndex        =   30
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label11 
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
            Left            =   1320
            TabIndex        =   29
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label5 
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
            Left            =   1320
            TabIndex        =   28
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Completion Date:"
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
            Left            =   1320
            TabIndex        =   27
            Top             =   1800
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Left            =   -73920
         TabIndex        =   1
         Top             =   600
         Width           =   7095
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   195
            Left            =   1680
            TabIndex        =   92
            Top             =   6720
            Width           =   1215
         End
         Begin VB.TextBox txtMS_PPT 
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
            Left            =   3840
            TabIndex        =   14
            Top             =   4200
            Width           =   975
         End
         Begin VB.TextBox txtMS_Xls 
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
            Left            =   3840
            TabIndex        =   13
            Top             =   3720
            Width           =   975
         End
         Begin VB.TextBox txtWord_Proc 
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
            Left            =   3840
            TabIndex        =   12
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox txtMS_Win 
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
            Left            =   3840
            TabIndex        =   11
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox txtApprec 
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
            Left            =   3840
            TabIndex        =   10
            Top             =   2280
            Width           =   975
         End
         Begin MSMask.MaskEdBox mskComp_Date 
            Height          =   270
            Left            =   2760
            TabIndex        =   9
            Top             =   1800
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   476
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
         Begin VB.Label lblOtherName 
            BackColor       =   &H8000000E&
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
            Left            =   4320
            TabIndex        =   112
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblClass_Result 
            BackColor       =   &H8000000E&
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
            Left            =   2760
            TabIndex        =   25
            Top             =   5640
            Width           =   2895
         End
         Begin VB.Label lblAverage_Score 
            BackColor       =   &H8000000E&
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
            Left            =   2760
            TabIndex        =   24
            Top             =   5160
            Width           =   2895
         End
         Begin VB.Label lblTotal_Score 
            BackColor       =   &H8000000E&
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
            Left            =   2760
            TabIndex        =   23
            Top             =   4680
            Width           =   2895
         End
         Begin VB.Label Label22 
            Caption         =   "Classification:"
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
            Left            =   960
            TabIndex        =   22
            Top             =   5760
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "Average Score:"
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
            Left            =   960
            TabIndex        =   21
            Top             =   5280
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Total Score:"
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
            Left            =   960
            TabIndex        =   20
            Top             =   4800
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "MS PowerPoint:"
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
            Left            =   2040
            TabIndex        =   19
            Top             =   4200
            Width           =   1935
         End
         Begin VB.Label Label10 
            Caption         =   "MS Excel:"
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
            Left            =   2040
            TabIndex        =   18
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "Word Processing:"
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
            Left            =   2040
            TabIndex        =   17
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "MS Windows:"
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
            Left            =   2040
            TabIndex        =   16
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label7 
            Caption         =   "Appreciation:"
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
            Left            =   2040
            TabIndex        =   15
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label lbl_Program 
            BackColor       =   &H8000000E&
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
            Left            =   2760
            TabIndex        =   8
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label lblStud_Name 
            BackColor       =   &H8000000E&
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
            Left            =   2760
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblAdm_No 
            BackColor       =   &H8000000E&
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
            Left            =   2760
            TabIndex        =   6
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label6 
            Caption         =   "Completion Date:"
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
            Left            =   960
            TabIndex        =   5
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label frmDiploma 
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
            Left            =   960
            TabIndex        =   4
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label2 
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
            Left            =   960
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Programme:"
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
            Left            =   960
            TabIndex        =   2
            Top             =   1320
            Width           =   1815
         End
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1920
         TabIndex        =   127
         Top             =   7320
         Width           =   5175
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   -73080
         TabIndex        =   123
         Top             =   7320
         Width           =   5175
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   -72960
         TabIndex        =   115
         Top             =   7320
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim SMS_DB As Database
 Dim rstRegist As Recordset
 Dim rstCert As Recordset
 Dim rstDip As Recordset
 Dim rstEngi As Recordset
Private Sub Command1_Click()
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
   lblAdm_No = !Adm_No
    lblStud_Name = !SurName
    lblOtherName = ![Other Name]
    lbl_Program = !Program
    End If
    End With
    mskComp_Date.SetFocus
    Call ClearCertForm
End Sub

Private Sub Command10_Click()
On Error Resume Next
With rstDip
ComputeRec
.Edit
!SurName = lblDSName
![Other Name] = lblDOther
![Reg No] = lblDANum
!Programme = lblDProg
![Completion Date] = mskDComp_Date
!Appreciation = txtDAppre
![MS Windows] = txtDWin
![Word Processing] = txtDWord
![MS Excel] = txtDExcel
![MS PowerPoint] = txtDPPt
!DBMS = txtDDBMS
![CorelDraw] = txtDCorelDraw
![VB Introduction] = txtDVB
!Seminar = txtDSeminar
![MS Access] = txtDAccess
![Total Score] = lblDTotal
![Average Score] = lblDAverage
!Classification = lblDClass
.Update
.Bookmark = .LastModified
End With
Call ClearDipform
End Sub

Private Sub Command11_Click()
On Error Resume Next
If MsgBox("Delete this Record?", vbYesNo + vbQuestion, "Delete") = vbNo Then
    Exit Sub
    End If
    With rstDip
    .Delete
    Call ClearDipform
    End With
End Sub

Private Sub Command12_Click()
On Error Resume Next
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
   lblEAdm = !Adm_No
   lblESname = !SurName
   lblEOtherName = ![Other Name]
   lblEProgram = !Program
    End If
    End With
    mskComp_Date.SetFocus
End Sub

Private Sub Command13_Click()
On Error Resume Next
Call ComputeRecE
Call SaveResultEng
Call ClearEngform
End Sub
Private Sub ComputeRecE()
On Error Resume Next
Dim A As Single, B As Single, C As Single, D As Single, E As Single
Dim F As Single, G As Single, H As Single, i As Single, J As Single
Dim K As Double, L As Double
A = CSng(txtEAppre.Text)
B = CSng(txtEWin.Text)
C = CSng(txtEDos.Text)
D = CSng(txtEHardIDN.Text)
E = CSng(txtEHardIN.Text)
F = CSng(txtESoftIN.Text)
G = CSng(txtETroubSh.Text)
H = CSng(txtEMaint.Text)
i = CSng(txtESystUp.Text)
J = CSng(txtEProj.Text)
K = A + B + C + D + E + F + G + H + i + J
lblETotal = Format(CDbl(K), "0.00")
L = K / 10
lblEAverage = Format(CDbl(L), "0.00")

If txtEAppre.Text < 40 Then
MsgBox "You have a Resit Paper in Computer Appreciation", vbInformation, "Resit"
lblETotal = ""
lblEAverage = ""
lblEClass = ""
End If
If txtEWin < 40 Then
MsgBox "You have a Resit Paper in MS Window", vbInformation, "Resit"
lblETotal = ""
lblEAverage = ""
lblEClass = ""
End If
If txtEDos < 40 Then
MsgBox "You have a Resit Paper in MS Dos", vbInformation, "Resit"
lblETotal = ""
lblEAverage = ""
lblEClass = ""
End If
If txtEHardIDN < 40 Then
MsgBox "You have a Resit Paper in Hardware Identification", vbInformation, "Resit"
lblETotal = ""
lblEAverage = ""
lblEClass = ""
End If
If txtEHardIN < 40 Then
MsgBox "You have a Resit Paper in Hardware Installation", vbInformation, "Resit"
lblETotal = ""
lblEAverage = ""
lblEClass = ""
End If
If txtESoftIN < 40 Then
MsgBox "You have a Resit Paper in Software Installation", vbInformation, "Resit"
lblETotal = ""
lblEAverage = ""
lblEClass = ""
End If
If txtETroubSh < 40 Then
MsgBox "You have a Resit Paper in Trouble Shooting", vbInformation, "Resit"
lblETotal = ""
lblEAverage = ""
lblEClass = ""
End If
If txtEMaint < 40 Then
MsgBox "You have a Resit Paper in Maintenance", vbInformation, "Resit"
lblETotal = ""
lblEAverage = ""
lblEClass = ""
End If
If txtESystUp < 40 Then
MsgBox "You have a Resit Paper in System Upgrading", vbInformation, "Resit"
lblETotal = ""
lblEAverage = ""
lblEClass = ""
End If

If txtEProj < 40 Then
MsgBox "You have a Resit Paper in Program", vbInformation, "Resit"
lblETotal = ""
lblEAverage = ""
lblEClass = ""
End If

Select Case lblEAverage
Case Is >= 70
lblEClass = "Distinction"
Case Is >= 60
lblEClass = "Upper Credit"
Case Is >= 50
lblEClass = "Lower Credit"
Case Is >= 40
lblEClass = "Pass"
Case Else
lblEClass = "Fail"
End Select
End Sub

Private Sub SaveResultEng()
On Error GoTo ErrTrap
With rstEngi
.AddNew
![Sur Name] = lblESname
![Other Name] = lblEOtherName
![Reg No] = lblEAdm
!Programme = lblEProgram
![Completion Date] = mskEComp_Date
!Appreciation = txtEAppre
![MS Windows] = txtEWin
![MS Dos] = txtEDos
![Hardware Identification] = txtEHardIDN
![Software Installation] = txtESoftIN
![System Trouble Shooting] = txtETroubSh
![Hardware Installation] = txtEHardIN
 ![System Upgrading ] = txtESystUp
![General Maintenance] = txtEMaint
!Project = txtEProj
![Total Score] = lblETotal
![Average Score] = lblEAverage
!Classification = lblEClass
.Update
.Bookmark = .LastModified
End With
Exit Sub
ErrTrap:
MsgBox Err.Description, vbInformation, "Error"
End Sub

Private Sub Command14_Click()
Dim StrSearch As String
    StrSearch = InputBox("Adminssion Number:", "Find Student")
    On Error Resume Next
    With rstEngi
    .Index = "Reg No"
    .Seek "=", StrSearch
    If .NoMatch Then
    MsgBox "There is No Record", vbInformation, "Find Student"
    Exit Sub
    Else
 GetRec3
    End If
    End With
End Sub
Private Sub GetRec3()
With rstEngi
 lblESname = ![Sur Name]
 lblEOtherName = ![Other Name]
 lblEAdm = ![Reg No]
 lblEProgram = !Programme
 mskEComp_Date = ![Completion Date]
 txtEAppre = !Appreciation
 txtEWin = ![MS Windows]
 txtEDos = ![MS Dos]
 txtEHardIDN = ![Hardware Identification]
 txtESoftIN = ![Software Installation]
 txtETroubSh = ![System Trouble Shooting]
 txtEHardIN = ![Hardware Installation]
 txtESystUp = ![System Upgrading ]
 txtEMaint = ![General Maintenance]
 txtEProj = !Project
 lblETotal = ![Total Score]
 lblEAverage = ![Average Score]
 lblEClass = !Classification
End With
End Sub
Private Sub Command15_Click()
On Error Resume Next
ComputeRecE
With rstEngi
.Edit
![Sur Name] = lblESname
![Other Name] = lblEOtherName
![Reg No] = lblEAdm
!Programme = lblEProgram
![Completion Date] = mskEComp_Date
!Appreciation = txtEAppre
![MS Windows] = txtEWin
![MS Dos] = txtEDos
![Hardware Identification] = txtEHardIDN
![Software Installation] = txtESoftIN
![System Trouble Shooting] = txtETroubSh
![Hardware Installation] = txtEHardIN
 ![System Upgrading ] = txtESystUp
![General Maintenance] = txtEMaint
!Project = txtEProj
![Total Score] = lblETotal
![Average Score] = lblEAverage
!Classification = lblEClass
.Update
.Bookmark = .LastModified
End With
Call ClearEngform
End Sub

Private Sub Command16_Click()
On Error Resume Next
If MsgBox("Delete this Record?", vbYesNo + vbQuestion, "Delete") = vbNo Then
    Exit Sub
    End If
    With rstEngi
    .Delete
    Call ClearEngform
    End With
End Sub

Private Sub Command17_Click()
On Error Resume Next
With rstCert
.MovePrevious
    If .BOF Then
    .MoveFirst
    End If
GetResult1
End With
End Sub

Private Sub Command18_Click()
On Error Resume Next
With rstCert
.MoveFirst
GetResult1
End With
End Sub

Private Sub Command19_Click()
On Error Resume Next
With rstCert
.MoveNext
    If .EOF Then
    .MoveLast
    End If
GetResult1
End With
End Sub

Private Sub Command2_Click()
On Error Resume Next
Call Compute
Call SaveResult
lblAdm_No = ""
    lblStud_Name = ""
    lblOtherName = ""
    lbl_Program = ""
    mskComp_Date = ""
Call ClearCertForm
End Sub
Private Sub Compute()
On Error Resume Next
Dim A As Single, B As Single, C As Single, D As Single, E As Single, F As Double, G As Double
A = CSng(txtApprec.Text)
B = CSng(txtMS_Win.Text)
C = CSng(txtWord_Proc.Text)
D = CSng(txtMS_PPT)
E = CSng(txtMS_Xls.Text)
F = A + B + C + D + E
lblTotal_Score = Format(CDbl(F), "0.00")
G = F / 5
lblAverage_Score = Format(CDbl(G), "0.00")
If txtApprec.Text < 40 Then
MsgBox "You have a Resit Paper in Computer Appreciation", vbInformation, "Resit"
lblTotal_Score = ""
lblAverage_Score = ""
lblClass_Result = ""
End If
If txtMS_Win.Text < 40 Then
MsgBox "You have a Resit Paper in MS Windows", vbInformation, "Resit"
lblTotal_Score = ""
lblAverage_Score = ""
lblClass_Result = ""
End If
If txtWord_Proc.Text < 40 Then
MsgBox "You have a Resit Paper in Word Processing", vbInformation, "Resit"
lblTotal_Score = ""
lblAverage_Score = ""
lblClass_Result = ""
End If
If txtMS_PPT.Text < 40 Then
MsgBox "You have a Resit Paper in MS Powerpoint", vbInformation, "Resit"
lblTotal_Score = ""
lblAverage_Score = ""
lblClass_Result = ""
End If
If txtMS_Xls.Text < 40 Then
MsgBox "You have a Resit Paper in MS Excel", vbInformation, "Resit"
lblTotal_Score = ""
lblAverage_Score = ""
lblClass_Result = ""
End If

Select Case lblAverage_Score
Case Is >= 70
lblClass_Result = "Distinction"
Case Is >= 60
lblClass_Result = "Upper Credit"
Case Is >= 50
lblClass_Result = "Lower Credit"
Case Is >= 40
lblClass_Result = "Pass"
Case Else
lblClass_Result = "Fail"
End Select

End Sub
Private Sub SaveResult()
On Error GoTo ErrTrap
With rstCert
.AddNew
![Sur Name] = lblStud_Name
![Other Name] = lblOtherName
![Reg No] = lblAdm_No
!Programme = lbl_Program
![Completion Date] = mskComp_Date
!Appreciation = txtApprec
![MS Windows] = txtMS_Win
![MS Word] = txtWord_Proc
![MS Excel] = txtMS_Xls
![MS PowerPoint] = txtMS_PPT
![Total Score] = lblTotal_Score
![Average Score] = lblAverage_Score
!Classification = lblClass_Result
.Update
.Bookmark = .LastModified
End With
Exit Sub
ErrTrap:
MsgBox Err.Description, vbInformation, "Error"

End Sub

Private Sub Command20_Click()
On Error Resume Next
With rstCert
.MoveLast
GetResult1
End With
End Sub

Private Sub Command21_Click()
On Error Resume Next
With rstDip
.MoveLast
GetRec2
End With
End Sub

Private Sub Command22_Click()
On Error Resume Next
With rstDip
.MoveNext
    If .EOF Then
    .MoveLast
    End If
GetRec2
End With
End Sub

Private Sub Command23_Click()
On Error Resume Next
With rstDip
.MoveFirst
GetRec2
End With
End Sub

Private Sub Command24_Click()
On Error Resume Next
With rstEngi
.MoveLast
GetRec3
End With
End Sub

Private Sub Command25_Click()
On Error Resume Next
With rstEngi
.MoveNext
    If .EOF Then
    .MoveLast
    End If
GetRec3
End With
End Sub

Private Sub Command26_Click()
On Error Resume Next
With rstEngi
.MoveFirst
GetRec3
End With
End Sub

Private Sub Command27_Click()
On Error Resume Next
With rstCert
.MovePrevious
    If .BOF Then
    .MoveFirst
    End If
GetRec2
End With
End Sub

Private Sub Command28_Click()
On Error Resume Next
With rstEngi
.MovePrevious
    If .BOF Then
    .MoveFirst
    End If
GetRec3
End With
End Sub

Private Sub Command4_Click()
Dim StrSearch As String
    StrSearch = InputBox("Adminssion Number:", "Find Student")
    On Error Resume Next
    With rstCert
    .Index = "Reg No"
    .Seek "=", StrSearch
    If .NoMatch Then
    MsgBox "There is No Record", vbInformation, "Find Student"
    Exit Sub
    Else
   GetResult1
    End If
    End With
    
End Sub
Private Sub GetResult1()
On Error Resume Next
With rstCert
 lblStud_Name = ![Sur Name]
 lblOtherName = ![Other Name]
 lblAdm_No = ![Reg No]
 lbl_Program = !Programme
 mskComp_Date = ![Completion Date]
 txtApprec = !Appreciation
 txtMS_Win = ![MS Windows]
 txtWord_Proc = ![MS Word]
 txtMS_Xls = ![MS Excel]
 txtMS_PPT = ![MS PowerPoint]
 lblTotal_Score = ![Total Score]
 lblAverage_Score = ![Average Score]
 lblClass_Result = !Classification
End With
End Sub

Private Sub Command5_Click()
On Error Resume Next
Call Compute
With rstCert
.Edit
![Sur Name] = lblStud_Name
![Other Name] = lblOtherName
![Reg No] = lblAdm_No
!Programme = lbl_Program
![Completion Date] = mskComp_Date
!Appreciation = txtApprec
![MS Windows] = txtMS_Win
![MS Word] = txtWord_Proc
![MS Excel] = txtMS_Xls
![MS PowerPoint] = txtMS_PPT
![Total Score] = lblTotal_Score
![Average Score] = lblAverage_Score
!Classification = lblClass_Result
.Update
.Bookmark = .LastModified
End With
Call ClearCertForm
End Sub

Private Sub Command6_Click()
On Error Resume Next
 If MsgBox("Delete this Record?", vbYesNo + vbQuestion, "Delete") = vbNo Then
    Exit Sub
    End If
    With rstCert
    .Delete
    lblAdm_No = ""
    lblStud_Name = ""
    lblOtherName = ""
    lbl_Program = ""
    mskComp_Date = ""
    Call ClearCertForm
    End With
End Sub

Private Sub Command7_Click()
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
   lblDANum = !Adm_No
   lblDSName = !SurName
   lblDOther = ![Other Name]
   lblDProg = !Program
    End If
    End With
    mskComp_Date.SetFocus
    
End Sub
Private Sub ComputeRec()
On Error Resume Next
Dim A As Single, B As Single, C As Single, D As Single, E As Single
Dim F As Single, G As Single, H As Single, i As Single, J As Single
Dim K As Double, L As Double
A = CSng(txtDAppre.Text)
B = CSng(txtDWin.Text)
C = CSng(txtDWord.Text)
D = CSng(txtDPPt.Text)
E = CSng(txtDAccess.Text)
F = CSng(txtDExcel.Text)
G = CSng(txtDDBMS.Text)
H = CSng(txtDCorelDraw.Text)
i = CSng(txtDVB.Text)
J = CSng(txtDSeminar.Text)
K = A + B + C + D + E + F + G + H + i + J
lblDTotal = Format(CDbl(K), "0.00")
L = K / 10
lblDAverage = Format(CDbl(L), "0.00")

If txtDAppre.Text < 40 Then
MsgBox "You have a Resit Paper in Computer Appreciation", vbInformation, "Resit"
lblDTotal = ""
lblDAverage = ""
lblDClass = ""
End If
If txtDWin < 40 Then
MsgBox "You have a Resit Paper in MS Window", vbInformation, "Resit"
lblDTotal = ""
lblDAverage = ""
lblDClass = ""
End If
If txtDWord < 40 Then
MsgBox "You have a Resit Paper in MS Word", vbInformation, "Resit"
lblDTotal = ""
lblDAverage = ""
lblDClass = ""
End If
If txtDPPt < 40 Then
MsgBox "You have a Resit Paper in MS Powerpoint", vbInformation, "Resit"
lblDTotal = ""
lblDAverage = ""
lblDClass = ""
End If
If txtDExcel < 40 Then
MsgBox "You have a Resit Paper in MS Excel", vbInformation, "Resit"
lblDTotal = ""
lblDAverage = ""
lblDClass = ""
End If
If txtDAccess < 40 Then
MsgBox "You have a Resit Paper in MS Access", vbInformation, "Resit"
lblDTotal = ""
lblDAverage = ""
lblDClass = ""
End If
If txtDDBMS < 40 Then
MsgBox "You have a Resit Paper in DBMS", vbInformation, "Resit"
lblDTotal = ""
lblDAverage = ""
lblDClass = ""
End If
If txtDCorelDraw < 40 Then
MsgBox "You have a Resit Paper in CorelDraw", vbInformation, "Resit"
lblDTotal = ""
lblDAverage = ""
lblDClass = ""
End If
If txtDVB < 40 Then
MsgBox "You have a Resit Paper in Visual Basic", vbInformation, "Resit"
lblDTotal = ""
lblDAverage = ""
lblDClass = ""
End If

If txtDSeminar < 40 Then
MsgBox "You have a Resit Paper in Seminar", vbInformation, "Resit"
lblDTotal = ""
lblDAverage = ""
lblDClass = ""
End If

Select Case lblDAverage
Case Is >= 70
lblDClass = "Distinction"
Case Is >= 60
lblDClass = "Upper Credit"
Case Is >= 50
lblDClass = "Lower Credit"
Case Is >= 40
lblDClass = "Pass"
Case Else
lblDClass = "Fail"
End Select
End Sub
Private Sub Command8_Click()
On Error Resume Next
Call ComputeRec
Call SaveResultDip
Call ClearDipform
End Sub
Private Sub SaveResultDip()
On Error GoTo ErrTrap
On Error Resume Next
With rstDip
.AddNew
![Sur Name] = lblDSName
![Other Name] = lblDOther
![Reg No] = lblDANum
!Programme = lblDProg
![Completion Date] = mskDComp_Date
!Appreciation = txtDAppre
![MS Windows] = txtDWin
![Word Processing] = txtDWord
![MS Excel] = txtDExcel
![MS PowerPoint] = txtDPPt
!DBMS = txtDDBMS
![CorelDraw] = txtDCorelDraw
![VB Introduction] = txtDVB
!Seminar = txtDSeminar
![MS Access] = txtDAccess
![Total Score] = lblDTotal
![Average Score] = lblDAverage
!Classification = lblDClass
.Update
.Bookmark = .LastModified
End With
Exit Sub
ErrTrap:
MsgBox Err.Description, vbInformation, "Error"
End Sub

Private Sub Command9_Click()
Dim StrSearch As String
    StrSearch = InputBox("Adminssion Number:", "Find Student")
    On Error Resume Next
    With rstDip
    .Index = "Reg No"
    .Seek "=", StrSearch
    If .NoMatch Then
    MsgBox "There is No Record", vbInformation, "Find Student"
    Exit Sub
    Else
 GetRec2
    End If
    End With
End Sub
Private Sub GetRec2()
On Error Resume Next
With rstDip
lblDSName = ![Sur Name]
 lblDOther = ![Other Name]
 lblDANum = ![Reg No]
 lblDProg = !Programme
 mskDComp_Date = ![Completion Date]
 txtDAppre = !Appreciation
 txtDWin = ![MS Windows]
 txtDWord = ![Word Processing]
 txtDExcel = ![MS Excel]
 txtDPPt = ![MS PowerPoint]
 txtDDBMS = !DBMS
 txtDCorelDraw = ![CorelDraw]
 txtDVB = ![VB Introduction]
 txtDSeminar = !Seminar
 txtDAccess = ![MS Access]
 lblDTotal = ![Total Score]
 lblDAverage = ![Average Score]
 lblDClass = !Classification
End With
End Sub
Private Sub Form_Load()
Set SMS_DB = OpenDatabase(App.Path & "\SMS.mdb", False, False)
Set rstRegist = SMS_DB.OpenRecordset("Registration")
Set rstCert = SMS_DB.OpenRecordset("Certificate")
Set rstDip = SMS_DB.OpenRecordset("Diploma")
Set rstEngi = SMS_DB.OpenRecordset("Engineering")

frmResult.Caption = "Enter Result Score"
'txtApprec.Text = 0
'txtMS_Win.Text = 0
'txtWord_Proc.Text = 0
'txtMS_PPT = 0
'txtMS_Xls.Text = 0
'lblTotal_Score = 0
'lblAverage_Score = 0
'lblClass_Result = 0
End Sub

Private Sub ClearCertForm()
txtApprec.Text = ""
txtMS_Win.Text = ""
txtWord_Proc.Text = ""
txtMS_PPT = ""
txtMS_Xls.Text = ""
lblTotal_Score = ""
lblAverage_Score = ""
lblClass_Result = ""
End Sub
Private Sub ClearDipform()
lblDSName = ""
 lblDOther = ""
 lblDANum = ""
 lblDProg = ""
 mskDComp_Date = ""
 txtDAppre = ""
 txtDWin = ""
 txtDWord = ""
 txtDExcel = ""
 txtDPPt = ""
 txtDDBMS = ""
 txtDCorelDraw = ""
 txtDVB = ""
 txtDSeminar = ""
 txtDAccess = ""
 lblDTotal = ""
 lblDAverage = ""
 lblDClass = ""
End Sub

Private Sub ClearEngform()
lblESname = ""
 lblEOtherName = ""
 lblEAdm = ""
 lblEProgram = ""
 mskEComp_Date = ""
 txtEAppre = ""
 txtEWin = ""
 txtEDos = ""
 txtEHardIDN = ""
 txtESoftIN = ""
 txtETroubSh = ""
 txtEHardIN = ""
 txtESystUp = ""
 txtEMaint = ""
 txtEProj = ""
 lblETotal = ""
 lblEAverage = ""
 lblEClass = ""
End Sub


Private Sub SSTab1_DblClick()
Dim StrSearch As String
    StrSearch = InputBox("Adminssion Number:", "Find Student")
    On Error Resume Next
    With rstRegist
    .Index = "REg No"
    .Seek "=", StrSearch
    If .NoMatch Then
    MsgBox "There is No Record", vbInformation, "Find Student"
    Exit Sub
    Else
   lblDANum = !Adm_No
   lblDSName = !SurName
   lblDOther = ![Other Name]
   lblDProg = !Program
    End If
    End With
    mskComp_Date.SetFocus
End Sub
