VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmsubject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subjects"
   ClientHeight    =   5250
   ClientLeft      =   3045
   ClientTop       =   1935
   ClientWidth     =   6720
   ClipControls    =   0   'False
   Icon            =   "frmsubject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6720
   Begin VB.CommandButton Command6 
      Caption         =   ">>"
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<<"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   6615
      Begin MSDataGridLib.DataGrid dtgsub 
         Height          =   2895
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483624
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Search"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin MSDataListLib.DataCombo dcsubject 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Category"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Subject Name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmsubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset
Dim Rec1 As ADODB.Recordset



Private Sub Command1_Click()
On Error Resume Next
    rec.MoveFirst
    
    rec.AddNew
    rec!SubjectNames = UCase(Text2.Text)
    rec!Category = dcSubject.Text
    rec.UpdateBatch
    rec.Requery
    
    If (Err.Number = 3021 Or Err.Number = 0) Then
    main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
    Call modform.ClearTextBoxes(Me)
Else
    MsgBox "Subject already added....!"
End If
Text2.Text = ""
End Sub

Private Sub Command2_Click()
On Error Resume Next

    If MsgBox("Are you sure do you want to delete the current record?", vbYesNo, "Delete") = vbYes Then
        rec.Delete
        rec.UpdateBatch
        rec.Requery
        dtgCode.Refresh
    End If

End Sub

Private Sub Command3_Click()
On Error GoTo Errhand
    rec!SubjectNames = UCase(Text2.Text)
    rec!Category = dcSubject.Text
    rec.UpdateBatch
    main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
    Exit Sub
Errhand:
    MsgBox "YOU CAN NOT ENTER DUPLICATE VALUE"
End Sub

Private Sub Command4_Click()


Unload Me

End Sub

Private Sub Command5_Click()
On Error Resume Next
rec.MovePrevious
    If rec.BOF Then rec.MoveFirst
    Text2.Text = rec!SubjectNames
    dcSubject.Text = rec!Category
    
End Sub

Private Sub Command6_Click()
On Error Resume Next
rec.MoveNext
    If rec.EOF Then rec.MoveLast
    Text2.Text = rec!SubjectNames
    dcSubject.Text = rec!Category
End Sub

Private Sub Command7_Click()
Dim s As String
On Error Resume Next
s = InputBox("Enter Subject Name ", SEARCH)
If (s <> "") Then
rec.MoveFirst
    rec.Find "SubjectNames = '" & Trim(s) & "'"
    If rec.EOF Then
        MsgBox "Cannot Find"
    Else
        Text2.Text = rec!SubjectNames
        dcSubject.Text = rec!Category
    End If
End If
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "OL SUBJECTS")
    Set rec = openDB.OpenRecord("SELECT * FROM OLSUBJECT")
    Set Rec1 = openDB.OpenRecord("select distinct(Category) from OLSUBJECT")


    dcSubject.ListField = "Category"
    Set dcSubject.RowSource = Rec1
    dcSubject.Text = ""



    rec.MoveFirst
    Set dtgsub.DataSource = rec
    dtgsub.Columns(2).Locked = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Text2.SetFocus
rec.Close
Rec1.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

