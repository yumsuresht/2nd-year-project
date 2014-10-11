VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmgrade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   5580
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5730
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Medium"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   4215
      Begin VB.OptionButton Option3 
         Caption         =   "Sinhala"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "English"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tamil"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Dialog.frx":030A
      Left            =   120
      List            =   "Dialog.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   5535
      Begin MSDataGridLib.DataGrid dtgCode 
         Height          =   3015
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5318
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
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Number of Students"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Grade"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Division"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rec As ADODB.Recordset
Dim RecGrade As ADODB.Recordset


Dim opt As String


Private Sub CancelButton_Click()
Unload Me

End Sub

Private Sub Command1_Click()
On Error Resume Next

    If MsgBox("Are you sure do you want to delete the current record?", vbYesNo, "Delete") = vbYes Then
        rec.Delete
        rec.UpdateBatch
        rec.Requery
        dtgCode.Refresh
    End If

End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim s As String
Dim s1 As String

If Option1.Value = True Then
s1 = "TAMIL"
ElseIf Option2.Value = True Then
s1 = "ENGLISH"
Else
s1 = "SINHALA"
End If


If Not (Combo1.Text = "" Or Text2.Text = "") Then
s = Combo1.Text + " " + Text2.Text
 rec.MoveFirst
    rec.Find "ClassName = '" & Trim(s) & "'"
    If rec.EOF Then
        rec.AddNew
        rec!ClassName = UCase(s)
        rec!No_Of_Students = Val(Text1.Text)
        rec!Medium = s1
        rec.UpdateBatch
        dtgCode.Refresh
      main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
      Else
        MsgBox "Class already added....!"
    End If
Else
MsgBox "BLANK FIELD"
End If
End Sub



Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim st1, st2, i As Integer
    Call modform.FormSize(Me, "Grades and Divisions")
    Set rec = openDB.OpenRecord("SELECT * FROM CLASS ORDER BY CLASSNAME")
    rec.MoveFirst
    Set dtgCode.DataSource = rec
    dtgCode.Caption = "Details of Classes"
    dtgCode.HeadFont.Bold = True
    dtgCode.HeadFont.Size = 10
    dtgCode.Columns(0).Alignment = dbgCenter
    
    
    Set RecGrade = openDB.OpenRecord("select * from school")

    st1 = RecGrade!StartGrade
    st2 = RecGrade!EndGrade
    
    For i = st1 To st2
    Combo1.AddItem i
    Next i
    
    
    
    
    
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Text2.SetFocus
rec.Close
RecGrade.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub
