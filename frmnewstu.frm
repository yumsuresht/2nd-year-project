VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmnewstu 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Student Registration"
   ClientHeight    =   5595
   ClientLeft      =   2040
   ClientTop       =   1440
   ClientWidth     =   10080
   Icon            =   "frmnewstu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   10080
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9763
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "New Applications"
      TabPicture(0)   =   "frmnewstu.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(9)=   "Label9"
      Tab(0).Control(10)=   "Label10"
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(12)=   "Label29"
      Tab(0).Control(13)=   "Text8"
      Tab(0).Control(14)=   "Text1"
      Tab(0).Control(15)=   "Text2"
      Tab(0).Control(16)=   "Text3"
      Tab(0).Control(17)=   "Text5"
      Tab(0).Control(18)=   "Combo1"
      Tab(0).Control(19)=   "Text6"
      Tab(0).Control(20)=   "Combo2"
      Tab(0).Control(21)=   "Text4"
      Tab(0).Control(22)=   "Combo3"
      Tab(0).Control(23)=   "DTPicker2"
      Tab(0).Control(24)=   "Text7"
      Tab(0).Control(25)=   "Command2"
      Tab(0).Control(26)=   "Command3"
      Tab(0).Control(27)=   "Command5"
      Tab(0).Control(28)=   "Command12"
      Tab(0).Control(29)=   "Command14"
      Tab(0).Control(30)=   "Command1"
      Tab(0).Control(31)=   "Command4"
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Registration"
      TabPicture(1)   =   "frmnewstu.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label14"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label17"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label18"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label19"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label20"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label21"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label22"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label23"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label24"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label25"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label26"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label34"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "dpadmin"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "DTPicker1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtTel"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtFaJob"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Combo5"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtCity"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Combo6"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtstreet"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtFa"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtStuName"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtPreSch"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtadminNo"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Command6"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Command9"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Command10"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "dtctem"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "dtcclass"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).ControlCount=   32
      TabCaption(2)   =   "A/L - Registration"
      TabPicture(2)   =   "frmnewstu.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "Text9"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(3)=   "Frame4"
      Tab(2).Control(4)=   "dcaladmin"
      Tab(2).Control(5)=   "Command18"
      Tab(2).Control(6)=   "Command17"
      Tab(2).Control(7)=   "Command16"
      Tab(2).Control(8)=   "Command15"
      Tab(2).Control(9)=   "Label59"
      Tab(2).Control(10)=   "Label27"
      Tab(2).ControlCount=   11
      Begin VB.Frame Frame1 
         Caption         =   "A/L Starem and Class"
         Height          =   975
         Left            =   -74760
         TabIndex        =   106
         Top             =   2400
         Width           =   4575
         Begin MSDataListLib.DataCombo dcalclass 
            Height          =   315
            Left            =   3000
            TabIndex        =   107
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcalstr 
            Height          =   315
            Left            =   720
            TabIndex        =   108
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label28 
            Caption         =   "A/L Streams"
            Height          =   255
            Left            =   720
            TabIndex        =   110
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label33 
            Caption         =   "A/L Class"
            Height          =   255
            Left            =   3000
            TabIndex        =   109
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Frame Frame3 
         Caption         =   "O/L Result- 1st Time"
         Height          =   5055
         Left            =   -70080
         TabIndex        =   78
         Top             =   360
         Visible         =   0   'False
         Width           =   4935
         Begin VB.CommandButton Command7 
            Caption         =   "Go 2nd Time"
            Height          =   315
            Left            =   3240
            TabIndex        =   105
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label Label58 
            Caption         =   "YEAR"
            Height          =   255
            Left            =   3120
            TabIndex        =   102
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label57 
            Caption         =   "INDEX NO"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label56 
            Caption         =   "YEAR"
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
            TabIndex        =   100
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label55 
            Caption         =   "INDEX NO"
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
            TabIndex        =   99
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label54 
            Caption         =   "ENGLISH"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label53 
            Caption         =   "SOCIAL STUDIES"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label52 
            Caption         =   "SCIENCE AND TECHNOLOGY"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label51 
            Caption         =   "MATHAMATICS"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label50 
            Caption         =   "LANGUAGE"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   2520
            Width           =   3855
         End
         Begin VB.Label Label49 
            Caption         =   "RELIGION"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   2880
            Width           =   3855
         End
         Begin VB.Label Label48 
            Caption         =   "AESTHETIC"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   3240
            Width           =   3855
         End
         Begin VB.Label Label47 
            Caption         =   "TECHNICAL"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   3600
            Width           =   3855
         End
         Begin VB.Label Label46 
            Caption         =   "ADITIONAL 1"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   3960
            Width           =   3855
         End
         Begin VB.Label Label45 
            Caption         =   "ADITIONAL 2"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   4320
            Width           =   3855
         End
         Begin VB.Label Label44 
            Caption         =   "SUB"
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
            TabIndex        =   88
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label35 
            Caption         =   "SUB"
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
            TabIndex        =   87
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label36 
            Caption         =   "SUB"
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
            TabIndex        =   86
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label37 
            Caption         =   "SUB"
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
            TabIndex        =   85
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label38 
            Caption         =   "SUB"
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
            TabIndex        =   84
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label39 
            Caption         =   "SUB"
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
            TabIndex        =   83
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label40 
            Caption         =   "SUB"
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
            TabIndex        =   82
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label Label41 
            Caption         =   "SUB"
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
            TabIndex        =   81
            Top             =   3600
            Width           =   495
         End
         Begin VB.Label Label42 
            Caption         =   "SUB"
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
            TabIndex        =   80
            Top             =   3960
            Width           =   495
         End
         Begin VB.Label Label43 
            Caption         =   "SUB"
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
            TabIndex        =   79
            Top             =   4320
            Width           =   495
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000000&
            BorderWidth     =   2
            X1              =   0
            X2              =   4680
            Y1              =   840
            Y2              =   840
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "A/L Selecting Subjects"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   69
         Top             =   3480
         Width           =   4575
         Begin MSDataListLib.DataCombo dcsub3 
            Height          =   315
            Left            =   2040
            TabIndex        =   72
            Top             =   1320
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcsub2 
            Height          =   315
            Left            =   2040
            TabIndex        =   71
            Top             =   840
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcsub1 
            Height          =   315
            Left            =   2040
            TabIndex        =   70
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label32 
            Caption         =   "Subject 3"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label31 
            Caption         =   "Subject 2"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label30 
            Caption         =   "Subject 1"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   360
            Width           =   855
         End
      End
      Begin MSDataListLib.DataCombo dtcclass 
         Height          =   315
         Left            =   2280
         TabIndex        =   68
         Top             =   3240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtctem 
         Height          =   315
         Left            =   240
         TabIndex        =   67
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Add"
         Height          =   375
         Left            =   -73920
         TabIndex        =   66
         Top             =   4320
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dcaladmin 
         Height          =   315
         Left            =   -74760
         TabIndex        =   65
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   -71280
         TabIndex        =   64
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Search"
         Height          =   375
         Left            =   -71280
         TabIndex        =   63
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Edit"
         Height          =   375
         Left            =   -71280
         TabIndex        =   62
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Register"
         Height          =   375
         Left            =   -71280
         TabIndex        =   61
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   8640
         TabIndex        =   60
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Height          =   375
         Left            =   2520
         Picture         =   "frmnewstu.frx":035E
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Search Temp Student"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Register"
         Height          =   375
         Left            =   7440
         TabIndex        =   58
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtadminNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   56
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtPreSch 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   3960
         Width           =   9495
      End
      Begin VB.TextBox txtStuName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   1320
         Width           =   5895
      End
      Begin VB.TextBox txtFa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   39
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtstreet 
         Appearance      =   0  'Flat
         Height          =   1245
         Left            =   6480
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox Combo6 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmnewstu.frx":07A0
         Left            =   240
         List            =   "frmnewstu.frx":07AA
         TabIndex        =   37
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6480
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   2520
         Width           =   3255
      End
      Begin VB.ComboBox Combo5 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmnewstu.frx":07C0
         Left            =   3840
         List            =   "frmnewstu.frx":07D9
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtFaJob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   34
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtTel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   33
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         Height          =   375
         Left            =   -67920
         TabIndex        =   32
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   -66720
         TabIndex        =   31
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Search"
         Height          =   375
         Left            =   -69120
         TabIndex        =   30
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Edit"
         Height          =   375
         Left            =   -70320
         TabIndex        =   29
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -71520
         TabIndex        =   28
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         Height          =   375
         Left            =   -72720
         TabIndex        =   27
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74760
         TabIndex        =   9
         Top             =   2760
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   -71160
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd - MMM - yyyy"
         Format          =   75497475
         CurrentDate     =   38211
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmnewstu.frx":07F5
         Left            =   -72360
         List            =   "frmnewstu.frx":07F7
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71160
         TabIndex        =   8
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmnewstu.frx":07F9
         Left            =   -70920
         List            =   "frmnewstu.frx":07FB
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68520
         TabIndex        =   7
         Top             =   2640
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmnewstu.frx":07FD
         Left            =   -74760
         List            =   "frmnewstu.frx":080D
         TabIndex        =   5
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   1365
         Left            =   -68520
         TabIndex        =   6
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74760
         TabIndex        =   3
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72600
         TabIndex        =   2
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   3480
         Width           =   9615
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   3840
         TabIndex        =   55
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd - MMM - yyyy"
         Format          =   75497475
         CurrentDate     =   38211
      End
      Begin MSComCtl2.DTPicker dpadmin 
         Height          =   285
         Left            =   6480
         TabIndex        =   76
         Top             =   3240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd - MMM - yyyy"
         Format          =   75497475
         CurrentDate     =   38211
      End
      Begin VB.Label Label59 
         Caption         =   "Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   103
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "D_OF_Admin"
         Height          =   255
         Left            =   6480
         TabIndex        =   77
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Admission Number"
         Height          =   255
         Left            =   3600
         TabIndex        =   57
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label25 
         Caption         =   "City"
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
         Left            =   6480
         TabIndex        =   54
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label24 
         Caption         =   "Street"
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
         Left            =   6480
         TabIndex        =   53
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "Temp-Number"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "Name of Student"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label21 
         Caption         =   "Name of Father"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label20 
         Caption         =   "Date of Birth"
         Height          =   255
         Left            =   3840
         TabIndex        =   49
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Religion"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Address"
         Height          =   255
         Left            =   6480
         TabIndex        =   47
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Father's Job"
         Height          =   255
         Left            =   3840
         TabIndex        =   46
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label16 
         Caption         =   "Last Grade the student passed"
         Height          =   255
         Left            =   3840
         TabIndex        =   45
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label15 
         Caption         =   "Privious School and Address"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Class Admitted"
         Height          =   255
         Left            =   2280
         TabIndex        =   43
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Telephone No"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label29 
         Caption         =   "Telephone No"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "Admission Number"
         Height          =   255
         Left            =   -74760
         TabIndex        =   25
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Class Admitted"
         Height          =   255
         Left            =   -72360
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Privious School and Address"
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Last Grade the student passed"
         Height          =   255
         Left            =   -70920
         TabIndex        =   22
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Father's Job"
         Height          =   255
         Left            =   -71160
         TabIndex        =   21
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Address"
         Height          =   255
         Left            =   -68520
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Religion"
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Date of Birth"
         Height          =   255
         Left            =   -71160
         TabIndex        =   18
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Name of Father"
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Name of Student"
         Height          =   255
         Left            =   -72600
         TabIndex        =   16
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Temp-Number"
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Street"
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
         Left            =   -68520
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "City"
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
         Left            =   -68520
         TabIndex        =   13
         Top             =   2400
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmnewstu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecStuClass As ADODB.Recordset
Dim TempStu As ADODB.Recordset
Dim RecStu As ADODB.Recordset
Dim RecStuID As ADODB.Recordset
Dim RecTEAV As ADODB.Recordset
Dim RecYEAV As ADODB.Recordset
Dim RecAct As ADODB.Recordset
Dim RecStu5 As ADODB.Recordset
Dim RecStu6 As ADODB.Recordset
Dim RecStu7 As ADODB.Recordset
Dim RecStu8 As ADODB.Recordset
Dim RecStu9 As ADODB.Recordset
Dim RecTermAvg As ADODB.Recordset
Dim RecOlresult As ADODB.Recordset
Dim i As Integer


Private Sub Combo3_Click()
Combo2.clear
For i = 5 To Val(Combo3.Text) - 1
Combo2.AddItem i
Next i
End Sub

Private Sub Command1_Click()
    Call modform.ClearTextBoxes(Me)

End Sub

Private Sub Command10_Click()
Unload Me
End Sub

Private Sub Command12_Click()
On Error Resume Next
frmTemStuSearch.Show
End Sub

Private Sub Command14_Click()
Unload Me
End Sub

Private Sub Command15_Click()
On Error Resume Next

If (dcsub1.Text = dcsub2.Text) Or (dcsub2.Text = dcsub3.Text) Or (dcsub3.Text = dcsub1.Text) Then
MsgBox "You check your subjects"
Exit Sub
End If

RecStu9.AddNew
RecStu9!stuid = Trim(dcaladmin.Text)
RecStu9!StrName = Trim(dcalstr.Text)
RecStu9!Subject1 = dcsub1.Text
RecStu9!Subject2 = dcsub2.Text
RecStu9!Subject3 = dcsub3.Text
RecStu9.UpdateBatch
RecStu9.Requery




RecStu7!Curr_Class = dcalclass.Text
RecStu7.UpdateBatch
RecStu7.Requery

'dcaladmin.Text = ""


RecTermAvg.MoveFirst
RecTermAvg.Find "StuID = '" & Trim(dcaladmin.Text) & "'"
RecTermAvg!grade = dcalclass.Text
RecTermAvg.UpdateBatch
RecTermAvg.Requery



End Sub

Private Sub Command18_Click()
Unload Me

End Sub

Private Sub Command2_Click()
On Error Resume Next
 TempStu.MoveFirst
 RecStuID.MoveFirst
 Text1.Text = RecStuID!TEMIDS + 1
        TempStu.AddNew
        TempStu!TemID = RecStuID!TEMIDS + 1
        TempStu!StudentName = Text2.Text
        TempStu!FatherName = Text3.Text
        TempStu!D_Of_Birth = DTPicker2.Value
        TempStu!Religion = Combo1.Text
        TempStu!Street = Text5.Text
        TempStu!City = Text6.Text
        TempStu!FatherJob = Text4.Text
        TempStu!Telephone = Text7.Text
        TempStu!LastGrade = Combo2.Text
        TempStu!Pri_School_Add = Text8.Text
        TempStu!class = Combo3.Text
        TempStu.UpdateBatch
        TempStu.Requery
        
        RecStuID!TEMIDS = RecStuID!TEMIDS + 1
        RecStuID.UpdateBatch
    main.stbMain.Panels(1).Text = "Staus: RecStuord successfully saved"
    Call modform.ClearTextBoxes(Me)
Command2.Enabled = False
Command4.Enabled = True

End Sub

Private Sub Command3_Click()
On Error Resume Next
TempStu.Delete
TempStu.UpdateBatch
TempStu.Requery

End Sub

Private Sub Command4_Click()
On Error Resume Next
If (Val(RecStuID!TEMIDS) = 0) Then
MsgBox "You must initilized the TemID"
frminitial.Show
frminitial.Text1.SetFocus
frminitial.Text1.BackColor = &H80000018

Unload Me
Exit Sub
End If


Call modform.ClearTextBoxes(Me)
Text1.Text = Val(RecStuID!TEMIDS) + 1
Command2.Enabled = True
Command4.Enabled = False


End Sub

Private Sub Command5_Click()
On Error Resume Next
        TempStu!StudentName = Text2.Text
        TempStu!FatherName = Text3.Text
        TempStu!D_Of_Birth = DTPicker2.Value
        TempStu!Religion = Combo1.Text
        TempStu!Street = Text5.Text
        TempStu!City = Text6.Text
        TempStu!FatherJob = Text4.Text
        TempStu!Telephone = Text7.Text
        TempStu!LastGrade = Combo2.Text
        TempStu!Pri_School_Add = Text8.Text
        TempStu!class = Combo3.Text
        TempStu.UpdateBatch
        TempStu.Requery
        
        MsgBox "Record sucessfully edited"
End Sub

Private Sub Command6_Click()
On Error Resume Next
If (Val(RecStuID!STUIDS) = 0) Then
MsgBox "You must initilized the StudentID"
frminitial.Show
frminitial.Text2.SetFocus
frminitial.Text2.BackColor = &H80000018

Unload Me
Exit Sub
End If


If (Trim(dtcclass.Text) = "") Then
MsgBox "You must select the Current Class"
Exit Sub
End If

 RecStu.MoveFirst
 RecStuID.MoveFirst
        RecStu.AddNew
        RecStu!stuid = Val(RecStuID!STUIDS) + 1
        RecStu!StudentName = txtStuName.Text
        RecStu!FatherName = txtFa.Text
        RecStu!D_Of_Admin = dpadmin.Value
        RecStu!D_Of_Birth = DTPicker1.Value
        RecStu!Religion = Combo6.Text
        RecStu!Street = txtstreet.Text
        RecStu!City = txtCity.Text
        RecStu!FatherJob = txtFaJob.Text
        RecStu!Telephone = txtTel.Text
        RecStu!LastGrade = Combo5.Text
        RecStu!Pri_School_Add = txtPreSch.Text
        RecStu!AdminGrade = dtcclass.Text
        RecStu.UpdateBatch
        RecStu.Requery
 
TempStu.Delete
TempStu.UpdateBatch
TempStu.Requery

dtctem.Text = ""
dtctem.Refresh

RecAct.AddNew
RecAct!stuid = Trim(txtadminNo.Text)
RecAct!Curr_Class = Trim(dtcclass.Text)
RecAct.UpdateBatch
RecAct.Requery


RecTEAV.AddNew
RecTEAV!stuid = Trim(txtadminNo.Text)
RecTEAV!StudentName = Trim(txtStuName.Text)
RecTEAV!grade = Trim(dtcclass.Text)
RecTEAV.UpdateBatch
RecTEAV.Requery

RecYEAV.AddNew
RecYEAV!stuid = Trim(txtadminNo.Text)
RecYEAV.UpdateBatch
RecYEAV.Requery



RecStuID!STUIDS = RecStuID!STUIDS + 1
RecStuID.UpdateBatch

main.stbMain.Panels(1).Text = "Staus: RecStuord successfully saved"
Call modform.ClearTextBoxes(Me)
Form_Load
End Sub

Private Sub Command7_Click()
On Error Resume Next
If (Command7.Caption = "Go 1st Time") Then
Command7.Caption = "Go 2nd Time"
Frame3.Caption = "O/L Result- 1st Time"
RecOlresult.MoveLast
Call displayOlres(RecOlresult)
ElseIf (Command7.Caption = "Go 2nd Time") Then

Command7.Caption = "Go 1st Time"
Frame3.Caption = "O/L Result- 2nd Time"

RecOlresult.MoveFirst
Call displayOlres(RecOlresult)
End If
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Command9_Click()
On Error Resume Next
frmTemStuSearch.Show

End Sub

Private Sub dcaladmin_Change()
On Error Resume Next
Dim OL As String
OL = Trim(dcaladmin.Text)
If OL <> "" Then
    RecStu7.MoveFirst
    RecStu7.Find "StuID = '" & OL & "'"
    If RecStu7.EOF Then
        Text9.Text = ""
        Frame3.Visible = False
    Else
        Text9.Text = RecStu7!StudentName
        Frame3.Visible = True
    End If
Else
    Text9.Text = ""
    Frame3.Visible = False
End If
If (RecStu7.RecordCount = 1) Then
Command7.Visible = False
End If
Set RecOlresult = openDB.OpenRecord("select * from OLRESULT where StuID='" & OL & "'")
Call displayOlres(RecOlresult)

End Sub

Private Sub dcalstr_Change()
Dim st As String
st = Left(dcalstr.Text, 1)
Set RecStu5 = openDB.OpenRecord("select * from CLASS where ClassName like '12 " + st + "%'")


Set RecStu8 = openDB.OpenRecord("select * from ALSUBJECT WHERE STREAM='" & dcalstr.Text & "'")


dcalclass.ListField = "ClassName"
Set dcalclass.RowSource = RecStu5
dcalclass.Text = ""

dcsub1.ListField = "SubjectNames"
Set dcsub1.RowSource = RecStu8
dcsub1.Text = ""

dcsub2.ListField = "SubjectNames"
Set dcsub2.RowSource = RecStu8
dcsub2.Text = ""

dcsub3.ListField = "SubjectNames"
Set dcsub3.RowSource = RecStu8
dcsub3.Text = ""

End Sub

Private Sub dtctem_Change()
Dim TempID As String
    On Error Resume Next
    
    
    TempID = Trim(dtctem.Text)
    If TempID <> "" Then
        TempStu.MoveFirst
        TempStu.Find "TemID = '" & TempID & "'"
        If TempStu.EOF Then
            txtStuName.Text = ""
            txtFa.Text = ""
            DTPicker1.Value = 1 / 1 / 1900
            Combo6.Text = ""
            txtstreet.Text = ""
            txtCity.Text = ""
            txtFaJob.Text = ""
            txtTel.Text = ""
            Combo5.Text = ""
            txtPreSch.Text = ""
            dtcclass.Text = ""
            Command6.Enabled = False

        Else
            txtadminNo.Text = Val(RecStuID!STUIDS) + 1
            txtStuName.Text = TempStu!StudentName
            txtFa.Text = TempStu!FatherName
            DTPicker1.Value = TempStu!D_Of_Birth
            Combo6.Text = TempStu!Religion
            txtstreet.Text = TempStu!Street
            txtCity.Text = TempStu!City
            txtFaJob.Text = TempStu!FatherJob
            txtTel.Text = TempStu!Telephone
            Combo5.Text = TempStu!LastGrade
            txtPreSch.Text = TempStu!Pri_School_Add
            Command6.Enabled = True
       End If
    Else
       Command6.Enabled = False
    End If
     
End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "Student Registration")


Set RecStuID = openDB.OpenRecord("SELECT * FROM IDS")
Set TempStu = openDB.OpenRecord("SELECT * FROM TEMPSTUDENTS")
Set RecStuClass = openDB.OpenRecord("SELECT * FROM CLASS")
Set RecStu = openDB.OpenRecord("SELECT * FROM MAINSTUDENTS")
Set RecAct = openDB.OpenRecord("SELECT * FROM ACTIVESTUDENT")
Set RecTEAV = openDB.OpenRecord("SELECT * FROM TERMAVG")
Set RecYEAV = openDB.OpenRecord("SELECT * FROM YEARAVERAGE")
Set RecStu7 = openDB.OpenRecord("select * from ACTIVESTUDENT A,MAINSTUDENTS M WHERE M.StuID=A.StuID and A.CURR_CLASS IN('11 R1','11 R2')")
Set RecStu6 = openDB.OpenRecord("select * from STREAM")
Set RecStu9 = openDB.OpenRecord("select * from FOLLOWSTREAM")

Set RecTermAvg = openDB.OpenRecord("SELECT * FROM TERMAVG ")




dpadmin.Value = Date
DTPicker1.Value = Date

Command2.Enabled = False
Command6.Enabled = False

dtctem.ListField = "TemID"
Set dtctem.RowSource = TempStu
dtctem.Text = ""
    
dtcclass.ListField = "ClassName"
Set dtcclass.RowSource = RecStuClass
dtcclass.Text = ""

dcaladmin.ListField = "StuID"
Set dcaladmin.RowSource = RecStu7
dcaladmin.Text = ""

dcalstr.ListField = "StrName"
Set dcalstr.RowSource = RecStu6
dcalstr.Text = ""


For i = 6 To 13
Combo3.AddItem i
Next i

dpadmin.Value = Date


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

RecStuClass.Close
TempStu.Close
RecStu.Close
RecStuID.Close
RecTEAV.Close
RecYEAV.Close
RecAct.Close
RecStu5.Close
RecStu6.Close
RecStu7.Close
RecStu8.Close
RecStu9.Close
RecTermAvg.Close
RecOlresult.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub
Public Sub displayOlres(RecOlresult As ADODB.Recordset)
On Error Resume Next
Label55.Caption = RecOlresult.Fields(1)
Label56.Caption = RecOlresult.Fields(2)

Label44.Caption = RecOlresult.Fields(3)
Label35.Caption = RecOlresult.Fields(4)
Label36.Caption = RecOlresult.Fields(5)
Label37.Caption = RecOlresult.Fields(6)

Label50.Caption = UCase(Left(RecOlresult.Fields(7), Len(RecOlresult.Fields(7)) - 2))
Label38.Caption = Right(RecOlresult.Fields(7), 1)

Label49.Caption = UCase(Left(RecOlresult.Fields(8), Len(RecOlresult.Fields(8)) - 2))
Label39.Caption = Right(RecOlresult.Fields(8), 1)

Label48.Caption = UCase(Left(RecOlresult.Fields(9), Len(RecOlresult.Fields(9)) - 2))
Label40.Caption = Right(RecOlresult.Fields(9), 1)

Label47.Caption = UCase(Left(RecOlresult.Fields(10), Len(RecOlresult.Fields(10)) - 2))
Label41.Caption = Right(RecOlresult.Fields(10), 1)

If (Right(RecOlresult.Fields(11), 1) <> "") Then
Label46.Visible = True
Label42.Visible = True
Label46.Caption = UCase(Left(RecOlresult.Fields(11), Len(RecOlresult.Fields(11)) - 2))
Label42.Caption = Right(RecOlresult.Fields(11), 1)
Else
Label46.Visible = False
Label42.Visible = False
End If

If (Right(RecOlresult.Fields(12), 1) <> "") Then
Label45.Visible = True
Label43.Visible = True
Label45.Caption = UCase(Left(RecOlresult.Fields(12), Len(RecOlresult.Fields(12)) - 2))
Label43.Caption = Right(RecOlresult.Fields(12), 1)
Else
Label45.Visible = False
Label43.Visible = False
End If
End Sub


Public Sub disp(TempID As String)
On Error Resume Next
'Dim TempID As String
If TempID <> "" Then
        TempStu.MoveFirst
        TempStu.Find "TemID = '" & TempID & "'"
        If TempStu.EOF Then
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Combo1.Text = ""
            DTPicker2.Value = Format(Date, "dd/mm/yyyy")
            Text5.Text = ""
            Text6.Text = ""
            Text4.Text = ""
            Text7.Text = ""
            Combo2.Text = ""
            Text8.Text = ""
            Combo3.Text = ""
            MsgBox "Cannot find"
        Else
            Text1.Text = TempStu!TemID
            Text2.Text = TempStu!StudentName
            Text3.Text = TempStu!FatherName
            Combo1.Text = TempStu!Religion
            DTPicker2.Value = TempStu!D_Of_Birth
            Text5.Text = TempStu!Street
            Text6.Text = TempStu!City
            Text4.Text = TempStu!FatherJob
            Text7.Text = TempStu!Telephone
            Combo2.Text = TempStu!LastGrade
            Text8.Text = TempStu!Pri_School_Add
            Combo3.Text = TempStu!class
           
        End If
End If
End Sub


Public Sub disp1(s As String)
dtctem.Text = s

Call dtctem_Change
End Sub
