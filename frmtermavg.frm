VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmtermavg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   Icon            =   "frmtermavg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8385
   Visible         =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Operations"
      Height          =   6375
      Left            =   6480
      TabIndex        =   9
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         ToolTipText     =   "Exit the window"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   11
         ToolTipText     =   "Save the changes"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Change 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   10
         ToolTipText     =   "Edit the last term  averages"
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dtcclass 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label term 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Average"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Class"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Year"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   6255
      Begin MSDataGridLib.DataGrid dgterm 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Term averages"
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8281
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
   End
End
Attribute VB_Name = "frmtermavg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rec1 As ADODB.Recordset
Dim Rec2 As ADODB.Recordset
Dim rec As ADODB.Recordset
Dim Rec3 As ADODB.Recordset
Dim Rec4 As ADODB.Recordset


Dim s As Integer


Private Sub Change_Click()
frmLogin.Show
main.Enabled = False
End Sub

Private Sub Command1_Click()
On Error Resume Next

Rec3.UpdateBatch
Rec3.Requery
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub dtcclass_Change()
On Error Resume Next
If (frmLogin.LoginSucceeded = True) Then
Set Rec3 = openDB.OpenRecord("select * from TERMAVG T where T.Grade='" + Trim(dtcclass.Text) + "'")
Rec3.Requery
dgterm.Refresh
Frame2.Visible = True
Else
Set Rec3 = openDB.OpenRecord("select StuID,StudentName," + term.Caption + " from TERMAVG where Grade='" + Trim(dtcclass.Text) + "'")
Rec3.Requery
Frame2.Visible = True
End If

If (Rec3.RecordCount = 0) Then
Command1.Enabled = False
'Command3.Enabled = False
Change.Enabled = False
Else
Command1.Enabled = True
'Command3.Enabled = True
Change.Enabled = True
End If


Set dgterm.DataSource = Rec3
dgterm.Caption = "Student Term Average"
dgterm.HeadFont.Bold = True
dgterm.HeadFont.Size = 9
dgterm.Columns(0).Locked = True
dgterm.Columns(1).Locked = True
dgterm.Columns(2).Locked = True





       
End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "Term Average")
Text1.Text = Year(Now)
Set Rec1 = openDB.OpenRecord("SELECT * FROM CLASS")

dtcclass.ListField = "ClassName"
Set dtcclass.RowSource = Rec1
dtcclass.Text = ""

dtcclass.ZOrder
s = Val(Month(Now))
If (s <= 4) Then
term.Caption = "Term1"
ElseIf (s <= 8) Then
term.Caption = "Term2"
Else
term.Caption = "Term3"
End If



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
rec.Close
Rec1.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Public Sub viewterms()
        dtcclass_Change
        frmLogin.LoginSucceeded = False
        
        
End Sub

