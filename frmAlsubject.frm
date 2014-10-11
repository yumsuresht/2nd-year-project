VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAlsubject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8760
   Icon            =   "frmAlsubject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   8760
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   3615
      Begin MSDataGridLib.DataGrid dgsubject 
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4471
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
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin MSDataListLib.DataCombo dcStream 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Stream"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Subject Name"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAlsubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStream As ADODB.Recordset
Dim RecAlsubject As ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
RecAlsubject.AddNew
RecAlsubject!SubjectNames = (Trim(Text1.Text))
RecAlsubject!Stream = (Trim(dcStream.Text))
RecAlsubject.UpdateBatch
RecAlsubject.Requery
If (Err.Number <> 0) Then
MsgBox "Cannot insert"
End If


Form_Load


End Sub

Private Sub Command3_Click()
On Error Resume Next
RecAlsubject.Delete
RecAlsubject.UpdateBatch
RecAlsubject.Requery

End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "A/L Subjects")
    Set RecStream = openDB.OpenRecord("SELECT * FROM STREAM")
    Set RecAlsubject = openDB.OpenRecord("SELECT * FROM ALSUBJECT")

dcStream.ListField = "StrName"
Set dcStream.RowSource = RecStream

    RecAlsubject.MoveFirst
    Set dgsubject.DataSource = RecAlsubject

Text1.SetFocus

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Text1.SetFocus
RecStream.Close
RecAlsubject.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

