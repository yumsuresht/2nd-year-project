VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmstulea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Leavings"
   ClientHeight    =   7755
   ClientLeft      =   540
   ClientTop       =   1935
   ClientWidth     =   11115
   ClipControls    =   0   'False
   Icon            =   "frmstulea.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11115
   Begin VB.CommandButton Command5 
      Caption         =   "Check"
      Height          =   380
      Left            =   9600
      TabIndex        =   25
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   380
      Left            =   9600
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear"
      Height          =   380
      Left            =   9600
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Leave"
      Height          =   380
      Left            =   9600
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabCaption(0)   =   "Leaving Details"
      TabPicture(0)   =   "frmstulea.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label20"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label21"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label22"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dtplea"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtdoad"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtname"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtfather"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtdofb"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtreli"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtadd"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtgrade"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtpur"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dcadmin"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtdes"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "listclub"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Results"
      TabPicture(1)   =   "frmstulea.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      Begin MSComctlLib.ListView listclub 
         Height          =   1455
         Left            =   2520
         TabIndex        =   76
         Top             =   3120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Post Held"
            Object.Width           =   1596
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Club/Union"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Year"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.TextBox txtdes 
         Appearance      =   0  'Flat
         Height          =   1455
         Left            =   2520
         TabIndex        =   28
         Top             =   5880
         Width           =   6615
      End
      Begin MSDataListLib.DataCombo dcadmin 
         Height          =   315
         Left            =   2520
         TabIndex        =   26
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame Frame3 
         Caption         =   "A/L Result"
         Height          =   5895
         Left            =   -70440
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CommandButton cmd2nd 
            Caption         =   "2nd Time"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3360
            TabIndex        =   50
            Top             =   5400
            Width           =   975
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            BorderWidth     =   2
            X1              =   100
            X2              =   4440
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label lblzsco 
            Caption         =   "Z-SCORE"
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
            Left            =   2400
            TabIndex        =   49
            Top             =   4920
            Width           =   1335
         End
         Begin VB.Label Label31 
            Caption         =   "Z-SCORE"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   4920
            Width           =   855
         End
         Begin VB.Label lblisl 
            Caption         =   "ISLAND"
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
            Left            =   2400
            TabIndex        =   47
            Top             =   4560
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "ISLAND"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   4560
            Width           =   735
         End
         Begin VB.Label lbldis 
            Caption         =   "DISTRICT"
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
            Left            =   2400
            TabIndex        =   45
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label lblgen 
            Caption         =   "GENERAL"
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
            Left            =   2400
            TabIndex        =   44
            Top             =   3360
            Width           =   375
         End
         Begin VB.Label lbleng 
            Caption         =   "ENGLISH"
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
            Left            =   2400
            TabIndex        =   43
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label lblsubject3re 
            Caption         =   "SUBJECT3"
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
            Left            =   2400
            TabIndex        =   42
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label lblsubject2re 
            Caption         =   "SUBJECT2"
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
            Left            =   2400
            TabIndex        =   41
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label lblsubject1re 
            Caption         =   "SUBJECT1"
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
            Left            =   2400
            TabIndex        =   40
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label19 
            Caption         =   "DISTRICT RANK"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   4200
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "GENERAL KNOWLEDGE"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   3360
            Width           =   2295
         End
         Begin VB.Label Label17 
            Caption         =   "ENGLISH"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   3000
            Width           =   2295
         End
         Begin VB.Label lblsubject3 
            Caption         =   "SUBJECT3"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   2400
            Width           =   2295
         End
         Begin VB.Label lblsubject2 
            Caption         =   "SUBJECT2"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label lblsubject1 
            Caption         =   "SUBJECT1"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblyear 
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
            Left            =   1440
            TabIndex        =   33
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "YEAR"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblindex 
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
            Left            =   1440
            TabIndex        =   31
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "INDEX NO"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "O/L Result- 1st Time"
         Height          =   5895
         Left            =   -75000
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CommandButton Command1 
            Caption         =   "Go 2nd Time"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3000
            TabIndex        =   75
            Top             =   5400
            Width           =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000000&
            BorderWidth     =   2
            X1              =   80
            X2              =   4320
            Y1              =   1320
            Y2              =   1320
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
            Left            =   3480
            TabIndex        =   74
            Top             =   4920
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
            Left            =   3480
            TabIndex        =   73
            Top             =   4560
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
            Left            =   3480
            TabIndex        =   72
            Top             =   4200
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
            Left            =   3480
            TabIndex        =   71
            Top             =   3840
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
            Left            =   3480
            TabIndex        =   70
            Top             =   3480
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
            Left            =   3480
            TabIndex        =   69
            Top             =   3120
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
            Left            =   3480
            TabIndex        =   68
            Top             =   2760
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
            Left            =   3480
            TabIndex        =   67
            Top             =   2400
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
            Left            =   3480
            TabIndex        =   66
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label34 
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
            Left            =   3480
            TabIndex        =   65
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label33 
            Caption         =   "ADITIONAL 2"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   4920
            Width           =   3255
         End
         Begin VB.Label Label32 
            Caption         =   "ADITIONAL 1"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   4560
            Width           =   3255
         End
         Begin VB.Label Label30 
            Caption         =   "TECHNICAL"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   4200
            Width           =   3255
         End
         Begin VB.Label Label28 
            Caption         =   "AESTHETIC"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   3840
            Width           =   3255
         End
         Begin VB.Label Label27 
            Caption         =   "RELIGION"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   3480
            Width           =   3135
         End
         Begin VB.Label Label26 
            Caption         =   "LANGUAGE"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   3120
            Width           =   3255
         End
         Begin VB.Label Label25 
            Caption         =   "MATHAMATICS"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label24 
            Caption         =   "SCIENCE AND TECHNOLOGY"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label Label23 
            Caption         =   "SOCIAL STUDIES"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   2400
            Width           =   1935
         End
         Begin VB.Label Label16 
            Caption         =   "ENGLISH"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label15 
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
            TabIndex        =   54
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label14 
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
            Left            =   1320
            TabIndex        =   53
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "INDEX NO"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "YEAR"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.TextBox txtpur 
         Appearance      =   0  'Flat
         Height          =   1125
         Left            =   2520
         TabIndex        =   22
         Top             =   4680
         Width           =   6615
      End
      Begin VB.TextBox txtgrade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtadd 
         Appearance      =   0  'Flat
         Height          =   1245
         Left            =   2520
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtreli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtdofb 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtfather 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtdoad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtplea 
         Height          =   285
         Left            =   6840
         TabIndex        =   27
         Top             =   1800
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
         Format          =   75759619
         CurrentDate     =   38211
      End
      Begin VB.Label Label1 
         Caption         =   "Descriptions"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Posts held in School"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label21 
         Caption         =   "Purpose of seeking character certificate"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "Date of Leaving"
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Last Grade the student passed"
         Height          =   495
         Left            =   5400
         TabIndex        =   11
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Religion"
         Height          =   255
         Left            =   5400
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Date of Birth"
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Name of Father with initial"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Name of Student with initial"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Date of Admission"
         Height          =   255
         Left            =   5400
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Admission Number"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmstulea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStu As ADODB.Recordset
Dim RecStu1 As ADODB.Recordset
Dim RecOld As ADODB.Recordset
Dim RecLib As ADODB.Recordset
Dim RecLib1 As ADODB.Recordset
Dim RecAlresult As ADODB.Recordset
Dim RecOlresult As ADODB.Recordset

Dim Recclubmem As ADODB.Recordset
Dim RecClubPost1 As ADODB.Recordset
Dim RecClubPost2 As ADODB.Recordset


Dim msg As Boolean




Private Sub cmd2nd_Click()
On Error Resume Next
If (cmd2nd.Caption = "2nd Time") Then
cmd2nd.Caption = "1st Time"
Frame3.Caption = "A/L Result- 1st Time"
RecAlresult.MoveLast
Call displayAlres(RecAlresult)

Else
cmd2nd.Caption = "2nd Time"
Frame3.Caption = "A/L Result- 2nd Time"

RecAlresult.MoveFirst
Call displayAlres(RecAlresult)
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If (Command1.Caption = "Go 1st Time") Then
Command1.Caption = "Go 2nd Time"
Frame2.Caption = "O/L Result- 1st Time"
RecOlresult.MoveLast
Call displayOlres(RecOlresult)
ElseIf (Command1.Caption = "Go 2nd Time") Then

Command1.Caption = "Go 1st Time"
Frame2.Caption = "O/L Result- 2nd Time"

RecOlresult.MoveFirst
Call displayOlres(RecOlresult)
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If (msg = True) Then
RecOld.AddNew
RecOld!stuid = Trim(dcadmin.Text)
RecOld!D_Of_Leave = dtplea.Value
RecOld!LastGrade = txtgrade.Text
RecOld!Purpose_Of_Leave = txtpur.Text
RecOld!DESCRIPTIONS = txtDes.Text
RecOld.UpdateBatch
RecOld.Requery


RecStu1.Delete
RecStu1.UpdateBatch
RecStu1.Requery


RecLib1.MoveFirst
RecLib1.Find "SCID = '" & Trim(dcadmin.Text) & "'"
RecLib1.Delete
RecLib1.UpdateBatch
RecLib1.Requery
End If


Form_Load
dcadmin.Text = ""


End Sub

Private Sub Command4_Click()
Unload Me
End Sub



Private Sub Command5_Click()
On Error Resume Next
Dim stuid As String
Dim status As String
stuid = Trim(dcadmin.Text)
    
           If stuid <> "" Then
                RecLib.MoveFirst
                RecLib.Find "SCID = '" & stuid & "'"
                If RecLib.EOF Then
                    MsgBox "Library... Ok"
                    Command2.Enabled = True
                    msg = True
                    
                Else
                    status = RecLib!LendStatus
                    Command2.Enabled = False
                    If (status = "LEND") Then
                        MsgBox "You must return the library book"
                    ElseIf (status = "FINE") Then
                        MsgBox "You must pay library fine payment"
                    End If
                    
                End If
            Else
                Call modform.ClearTextBoxes(Me)
            End If

End Sub

Private Sub dcadmin_Change()
On Error Resume Next
Dim stuid As String
   Command2.Enabled = False
    stuid = Trim(dcadmin.Text)
    
           If stuid <> "" Then
                RecStu.MoveFirst
                RecStu.Find "StuID = '" & stuid & "'"
                RecStu1.Find "StuID = '" & stuid & "'"
                If RecStu.EOF Then
                Call modform.ClearTextBoxes(Me)
                Frame3.Visible = False
                Else
                    txtName.Text = RecStu!StudentName
                    txtfather.Text = RecStu!FatherName
                    txtadd.Text = RecStu!Street + "," + RecStu!City
                    txtdoad.Text = RecStu!D_Of_Admin
                    txtdofb.Text = RecStu!D_Of_Birth
                    txtreli.Text = RecStu!Religion
                    txtgrade.Text = Trim(Left(RecStu!Curr_Class, 2))
                    
                Set RecAlresult = openDB.OpenRecord("select * from ALRESULT A,FOLLOWSTREAM F WHERE A.STUID=F.STUID AND F.STUID='" & stuid & "' order by Alyear")
                Set RecOlresult = openDB.OpenRecord("select * from OLRESULT where StuID='" & stuid & "'")
                
                Set Recclubmem = openDB.OpenRecord("select * from CLUBMEMBER where STUID='" & stuid & "'")
                Set RecClubPost1 = openDB.OpenRecord("select CName,Years,Pres_StuID from CLUBMAINTAINCE where Pres_StuID='" & stuid & "'")
                Set RecClubPost2 = openDB.OpenRecord("select CName,Years,Sec_StuID from CLUBMAINTAINCE where Sec_StuID='" & stuid & "'")
                
                
                Call fillclubdetails1
                
                If (RecOlresult.RecordCount <> 0) Then
                    Frame2.Visible = True
                Else
                    Frame2.Visible = False
                End If
                
                
                If (RecAlresult.RecordCount <> 0) Then
                    Frame3.Visible = True
                Else
                    Frame3.Visible = False
                End If
                If (RecAlresult.RecordCount = 2) Then
                    cmd2nd.Enabled = True
                Else
                    cmd2nd.Enabled = False
                End If
                
                
                If (RecOlresult.RecordCount = 2) Then
                    Command1.Enabled = True
                Else
                    Command1.Enabled = False
                End If
                
                Call displayAlres(RecAlresult)
                Call displayOlres(RecOlresult)
                End If
            Else
                Call modform.ClearTextBoxes(Me)
                Frame3.Visible = False

            End If
            

End Sub

Private Sub Form_Load()
On Error Resume Next

Call modform.FormSize(Me, "STUDENT LEAVING")
    Set RecStu = openDB.OpenRecord("select *  from MAINSTUDENTS M,ACTIVESTUDENT A where M.StuID=A.StuID")
    Set RecStu1 = openDB.OpenRecord("select *  from ACTIVESTUDENT")
    Set RecOld = openDB.OpenRecord("select * from OLDBOYS")
    Set RecLib = openDB.OpenRecord("select * from LIBRARYMEMBER WHERE LENDSTATUS IN('LEND','FINE')")
    Set RecLib1 = openDB.OpenRecord("select * from LIBRARYMEMBER ")
    

    
    RecStu.MoveFirst
    dcadmin.ListField = "StuID"
    Set dcadmin.RowSource = RecStu
    
    dtplea.Value = Date
    Command2.Enabled = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecStu.Close
RecStu1.Close
RecOld.Close
RecLib.Close
RecLib1.Close
RecAlresult.Close
RecOlresult.Close
Recclubmem.Close
RecClubPost1.Close
RecClubPost2.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub


Public Sub displayAlres(RecAlresult As ADODB.Recordset)

lblindex.Caption = RecAlresult.Fields(2)
lblyear.Caption = RecAlresult.Fields(3)
                
lblsubject1.Caption = RecAlresult.Fields(16)
lblsubject1re.Caption = RecAlresult.Fields(5)
                
lblsubject2.Caption = RecAlresult.Fields(17)
lblsubject2re.Caption = RecAlresult.Fields(6)
                
lblsubject3.Caption = RecAlresult.Fields(18)
lblsubject3re.Caption = RecAlresult.Fields(7)
                
lbleng.Caption = RecAlresult.Fields(9)
lblgen.Caption = RecAlresult.Fields(10)
                
lbldis.Caption = RecAlresult.Fields(11)
lblisl.Caption = RecAlresult.Fields(12)
lblzsco.Caption = RecAlresult.Fields(13)
End Sub

Public Sub displayOlres(RecOlresult As ADODB.Recordset)
Label15.Caption = RecOlresult.Fields(1)
Label14.Caption = RecOlresult.Fields(2)

Label34.Caption = RecOlresult.Fields(3)
Label35.Caption = RecOlresult.Fields(4)
Label36.Caption = RecOlresult.Fields(5)
Label37.Caption = RecOlresult.Fields(6)

Label26.Caption = UCase(Left(RecOlresult.Fields(7), Len(RecOlresult.Fields(7)) - 2))
Label38.Caption = Right(RecOlresult.Fields(7), 1)

Label27.Caption = UCase(Left(RecOlresult.Fields(8), Len(RecOlresult.Fields(8)) - 2))
Label39.Caption = Right(RecOlresult.Fields(8), 1)

Label28.Caption = UCase(Left(RecOlresult.Fields(9), Len(RecOlresult.Fields(9)) - 2))
Label40.Caption = Right(RecOlresult.Fields(9), 1)

Label30.Caption = UCase(Left(RecOlresult.Fields(10), Len(RecOlresult.Fields(10)) - 2))
Label41.Caption = Right(RecOlresult.Fields(10), 1)

If (RecOlresult.Fields(11) <> "") Then
Label32.Visible = True
Label42.Visible = True
Label32.Caption = UCase(Left(RecOlresult.Fields(11), Len(RecOlresult.Fields(11)) - 2))
Label42.Caption = Right(RecOlresult.Fields(11), 1)
Else
Label32.Visible = False
Label42.Visible = False
End If

If (RecOlresult.Fields(12) <> "") Then
Label33.Visible = True
Label43.Visible = True
Label33.Caption = UCase(Left(RecOlresult.Fields(12), Len(RecOlresult.Fields(12)) - 2))
Label43.Caption = Right(RecOlresult.Fields(12), 1)
Else
Label33.Visible = False
Label43.Visible = False
End If
End Sub

Public Sub fillclubdetails1()
On Error Resume Next
If Not (Recclubmem.EOF And Recclubmem.BOF) Then
    Recclubmem.MoveFirst
        listclub.ListItems.clear
        While Not Recclubmem.EOF
            Set List = listclub.ListItems.Add
            List.Text = "Member"
            List.SubItems(1) = Recclubmem.Fields(0)
            List.SubItems(2) = " (" & Format(Recclubmem.Fields(2), "dd-mm-yyyy") & " to " & Format(Date, "dd-mm-yyyy") & " )"
            Recclubmem.MoveNext
        Wend
        Else
        listclub.ListItems.clear
    End If

End Sub
