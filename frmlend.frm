VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmlend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Lending"
   ClientHeight    =   8715
   ClientLeft      =   240
   ClientTop       =   1935
   ClientWidth     =   10425
   ClipControls    =   0   'False
   Icon            =   "frmlend.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10425
   Begin MSDataGridLib.DataGrid dgbook1 
      Height          =   4335
      Left            =   120
      TabIndex        =   55
      Top             =   4320
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7646
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame8 
      Caption         =   "Member Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   4980
      Begin VB.TextBox re3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox re2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo dcmrece 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label7 
         Caption         =   "Student ID"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Member ID"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Book Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4980
      Begin VB.CheckBox over 
         Caption         =   "Reference"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H80000011&
         Caption         =   "Search Book"
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dcblen 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label22 
         Caption         =   "Author"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "Category"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Book Name"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Access No"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label35 
         Caption         =   "Book ID"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   4215
      Left            =   45
      TabIndex        =   23
      Top             =   45
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lend"
      TabPicture(0)   =   "frmlend.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "Frame11"
      Tab(0).Control(2)=   "Frame13"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Reserve"
      TabPicture(1)   =   "frmlend.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame10 
         Caption         =   "Summery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   5160
         TabIndex        =   46
         Top             =   1560
         Width           =   3375
         Begin VB.Label Label38 
            Caption         =   "Access No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   54
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label39 
            Caption         =   "Member ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   53
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label40 
            Caption         =   "Access No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   52
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label41 
            Caption         =   "Member ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label36 
            Caption         =   "Reserved Date"
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
            Left            =   240
            TabIndex        =   50
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label37 
            Caption         =   "Last Reserve"
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
            Left            =   240
            TabIndex        =   49
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label45 
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   48
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label46 
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   47
            Top             =   1560
            Width           =   1455
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Reserve"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   5160
         TabIndex        =   41
         Top             =   360
         Width           =   3375
         Begin MSComCtl2.DTPicker redate 
            Height          =   375
            Left            =   1680
            TabIndex        =   42
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/M/yyyy"
            Format          =   45416451
            CurrentDate     =   38282
            MinDate         =   2
         End
         Begin VB.Label Label42 
            Caption         =   "Reserve Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label43 
            Caption         =   "Last Reserved Date"
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
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label 
            Caption         =   "Date"
            Height          =   255
            Left            =   2040
            TabIndex        =   43
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame14 
         Height          =   3735
         Left            =   8640
         TabIndex        =   36
         Top             =   360
         Width           =   1575
         Begin VB.CommandButton Command6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   1320
            Width           =   1100
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Clear"
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   840
            Width           =   1100
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Reserve"
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1100
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H80000011&
            Caption         =   "SEARCH"
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1800
            Width           =   1100
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Summery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -69840
         TabIndex        =   27
         Top             =   360
         Width           =   3375
         Begin VB.Label lblbid 
            Caption         =   "Access No"
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
            Left            =   1440
            TabIndex        =   35
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblmem 
            Caption         =   "Member ID"
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
            Left            =   1440
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Access No"
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
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Member ID"
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
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbldue 
            Caption         =   "Due Date"
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
            Left            =   1440
            TabIndex        =   31
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label lbldate 
            Caption         =   "Issue Date"
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
            Left            =   1440
            TabIndex        =   30
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label11 
            Caption         =   "Due Date"
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
            Left            =   120
            TabIndex        =   29
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Issue Date"
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
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Reserved Members"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -69840
         TabIndex        =   25
         Top             =   2640
         Width           =   3375
         Begin VB.CommandButton Command1 
            BackColor       =   &H80000011&
            Caption         =   "SEARCH"
            Height          =   375
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   360
            Width           =   1100
         End
      End
      Begin VB.Frame Frame13 
         Height          =   3735
         Left            =   -66360
         TabIndex        =   24
         Top             =   360
         Width           =   1575
         Begin VB.CommandButton Lend 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Lend"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1100
         End
         Begin VB.CommandButton clear 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Clear"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   840
            Width           =   1100
         End
         Begin VB.CommandButton exit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cancel"
            Height          =   374
            Left            =   240
            TabIndex        =   11
            Top             =   1320
            Width           =   1100
         End
         Begin VB.Timer Timer1 
            Interval        =   100
            Left            =   720
            Top             =   2400
         End
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2760
   End
End
Attribute VB_Name = "frmlend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecBook, RecMem, RecID, Recreser, RecLend, Recdate As ADODB.Recordset
Public Recbook2 As ADODB.Recordset
Dim da As Variant

Private Sub Command1_Click()
On Error Resume Next

    Recreser.MoveFirst
    Set frmlend.dgbook1.DataSource = Recreser
    frmlend.dgbook1.Caption = "Reserve Details"
    frmlend.dgbook1.HeadFont.Bold = True
    frmlend.dgbook1.HeadFont.Size = 10
    frmlend.dgbook1.Columns(0).Alignment = dbgCenter
   
End Sub

Private Sub Command10_Click()

End Sub

Private Sub Command2_Click()
On Error Resume Next

    Recreser.MoveFirst
    Set frmlend.dgbook1.DataSource = Recreser
    frmlend.dgbook1.Caption = "Reserve Details"
    frmlend.dgbook1.HeadFont.Bold = True
    frmlend.dgbook1.HeadFont.Size = 10
    
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim msg As String

If (Text16.Text = "") Then
MsgBox "Check your Book ID"
Exit Sub
End If

If (re2.Text = "") Then
MsgBox "Check your membership ID"
Exit Sub
End If

If (over.Value = 1) Then
    MsgBox "You can't borrow this book"
    Exit Sub
End If

If (dcblen.Text = "") Then
    MsgBox "Invalid Book"
    Exit Sub
ElseIf (dcmrece.Text = "") Then
    MsgBox "Invalid member"
    Exit Sub
End If

If (RecBook!BookStatus = "LEND") Then
msg = MsgBox("The book is borrowed, you can borrow this book after 7 days", vbOKCancel)
End If

If (msg = vbOK) Then
    redate.Value = Date + 7
    Recreser.AddNew
    Recreser!memid = Trim(dcmrece.Text)
    Recreser!AccessNo = dcblen.Text
    Recreser!ReserveDate = redate.Value
    Recreser!ResStatus = "NO"
    Recreser.UpdateBatch
    Recreser.Requery
        
    'Recbook!BookStatus = "RES"
    'Recbook.UpdateBatch
              
    'Recmem!LendStatus = "RES"
    'Recmem.UpdateBatch
    'Recmem.Requery
End If
        'MsgBox Err.Number
    If (Err.Number = 0 Or Err.Number = 13) Then
        main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
    Else
        MsgBox "You already reserve this book"
    End If
    redate.Value = Date
End Sub

Private Sub Command9_Click()
frmbooksearch.Show
End Sub

Private Sub dcblen_Change()

Dim bookno As String
    On Error Resume Next
    bookno = Trim(dcblen.Text)
    
               If bookno <> "" Then
                RecBook.MoveFirst
                RecBook.Find "AccessNo = '" & bookno & "'"
                If RecBook.EOF Then
                    Text17.Text = ""
                    Text16.Text = ""
                    Text15.Text = ""
                    Text14.Text = ""
                  
                Else
                    Text17.Text = RecBook!BookID
                    over.Value = RecBook!Overnight
                    Text16.Text = RecBook!title
                    Text15.Text = RecBook!Catagory
                    Text14.Text = RecBook!AuthorName
                    
                    
               End If
               Else
                    Text17.Text = ""
                    Text16.Text = ""
                    Text15.Text = ""
                    Text14.Text = ""
            End If
 


End Sub

Private Sub dcmrece_Change()
Dim memid As String
On Error Resume Next
    memid = Trim(dcmrece.Text)
    If memid <> "" Then
        RecMem.MoveFirst
        RecMem.Find "MemID = '" & memid & "'"
        If RecMem.EOF Then
            Label7.Caption = "Student ID"
            re2.Text = ""
            re3.Text = ""
            re4.Text = ""
        Else
            Label7.Caption = RecMem!status + " ID"
            re2.Text = RecMem!SCID
            re3.Text = RecMem!MemName
        End If
    Else
            re2.Text = ""
            re3.Text = ""
            re4.Text = ""
    
    End If
    
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call modform.FormSize(Me, "Library Transactions")
        Set RecBook = openDB.OpenRecord("select * from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID")
        Set RecMem = openDB.OpenRecord("select * from LIBRARYMEMBER where LendStatus IN('YES','RES','FINE')")
        Set RecID = openDB.OpenRecord("SELECT * FROM IDS")
        Set Recreser = openDB.OpenRecord("SELECT * FROM RESERVE")
        Set RecLend = openDB.OpenRecord("SELECT * FROM LENDING")
        

        dcblen.ListField = "AccessNo"
        Set dcblen.RowSource = RecBook
        
        dcmrece.ListField = "MemID"
        Set dcmrece.RowSource = RecMem
       redate.Value = Date
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecBook.Close
RecMem.Close
RecID.Close
Recreser.Close
RecLend.Close

Recbook2.Close
Recdate.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Lend_Click()
On Error Resume Next

If (Text16.Text = "") Then
MsgBox "Check your Book ID"
Exit Sub
End If

If (re2.Text = "") Then
MsgBox "Check your membership ID"
Exit Sub
End If





If (lbldue.Caption = Trim("you can not borrow")) Then
    MsgBox "You can not borrow this book"
    Exit Sub
End If

If (RecBook!BookStatus = "LEND") Then
    MsgBox "The book is borrowed"
    Exit Sub
End If

If (RecMem!LendStatus = "FINE") Then
    MsgBox "You must pay Fine"
    Exit Sub
End If


    RecLend.AddNew
    RecLend!TransNo = Val(RecID!LIBTRANSNO) + 1
    RecLend!memid = Trim(dcmrece.Text)
    RecLend!AccessNo = dcblen.Text
    RecLend!BorrowDate = Date
    RecLend!DueDate = da
    RecLend.UpdateBatch
    RecLend.Requery
        
    If (Err.Number <> 0) Then
        MsgBox "you can not borrow this book today"
        Exit Sub
    End If
        
        
    If (Recreser!AccessNo = Trim(dcblen.Text)) And (Recreser!ResStatus = "NO") And (Recreser!memid = Trim(dcmrece.Text)) Then
        Recreser.Delete
        Recreser.UpdateBatch
        Recreser.Requery
    ElseIf (Recreser!AccessNo = Trim(dcblen.Text)) And (Recreser!ResStatus = "NO") And (Recreser!memid <> Trim(dcmrece.Text)) Then
        MsgBox "The Book is Borrowed by another person"
        Exit Sub
    End If
        
        
        
    RecBook!BookStatus = "LEND"
    RecBook.UpdateBatch
              
    RecID!LIBTRANSNO = Val(RecID!LIBTRANSNO) + 1
    RecID.UpdateBatch
        
    RecMem!LendStatus = "LEND"
    RecMem.UpdateBatch
    RecMem.Requery
       
    Recbook2.Requery
    dcmrece.Text = ""
    
    main.stbMain.Panels(1).Text = "Staus: Record successfully saved"

End Sub

Private Sub Timer1_Timer()
On Error Resume Next

lblbid.Caption = dcblen.Text
lbldate.Caption = Date
lblmem.Caption = dcmrece.Text
    If (over.Value = 1) Then
        If ((Format(Now, "dddd") = "Friday") And (Label7.Caption = "Staff ID")) Then
            lbldue.Caption = (Date + 3) & " , " & Format(Now + 3, "dddd")
            da = Date + 3
        Else
            lbldue.Caption = "you can not borrow"
        End If
    Else
            da = Date + 7
            lbldue.Caption = (Date + 7) & " , " & Format(Now + 7, "dddd")
    End If

    Set Recdate = openDB.OpenRecord("select ReserveDate from RESERVE where AccessNo='" + dcblen.Text + "'")
Recdate.MoveLast
If (Err.Number = 3021) Then
Label.Caption = "Not Reserved"
Else
Label.Caption = Recdate!ReserveDate
End If
End Sub
