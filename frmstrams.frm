VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmstreams 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
   Icon            =   "frmstrams.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6810
   Begin VB.CommandButton Command7 
      Caption         =   ">>"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<<"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Search"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5175
      Begin MSDataGridLib.DataGrid dtgCode 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
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
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Stream ID"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmstreams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset
Dim Rec1 As ADODB.Recordset

Private Sub Command1_Click()
On Error Resume Next
If Not (Text1.Text = "" Or Text2.Text = "") Then
 rec.MoveFirst
 Rec1.MoveFirst
    rec.Find "ID = '" & Trim(Text1.Text) & "'"
    Rec1.Find "Name = '" & Trim(Text2.Text) & "'"
    If rec.EOF And Rec1.EOF Then
        rec.AddNew
        rec!id = UCase(Text1.Text)
        rec!name = UCase(Text2.Text)
        rec!StartYear = dtp.Value
        rec!DESCRIPTIONS = UCase(Text4.Text)
        
        rec.UpdateBatch
        dtgCode.Refresh
      main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
      Else
        MsgBox "Stream already added....!"
    End If
    Text1.Text = ""
    Text2.Text = ""
Else
MsgBox "BLANK FIELD"
End If
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
    rec!id = Text1.Text
    rec!name = Text2.Text
    rec.UpdateBatch
    main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
    Exit Sub
Errhand:
    MsgBox "YOU CAN NOT ENTER DUPLICATE VALUE"
End Sub

Private Sub Command4_Click()
Dim s As String
On Error Resume Next
s = InputBox("ENTER FIND VALUE", SEARCH)
If (s <> "") Then
rec.MoveFirst
 Rec1.MoveFirst
    rec.Find "ID = '" & Val(Trim(s)) & "'"
    Rec1.Find "Name = '" & UCase(Trim(s)) & "'"
    If rec.EOF Then
        If Rec1.EOF Then
        MsgBox "CAN NOT FIND"
        Else
        Text1.Text = Rec1!id
        Text2.Text = Rec1!name
        
        End If
    Else
        Text1.Text = rec!id
        Text2.Text = rec!name
              
    End If
End If
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
On Error Resume Next
rec.MovePrevious
    If rec.BOF Then rec.MoveFirst
    Text1.Text = rec!id
        Text2.Text = rec!name
        
End Sub

Private Sub Command7_Click()
On Error Resume Next
rec.MoveNext
    If rec.EOF Then rec.MoveLast
    Text1.Text = rec!id
        Text2.Text = rec!name
        
End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "A/L Streams")
    Set rec = openDB.OpenRecord("SELECT StrID AS ID,StrName AS Name FROM Stream")
    Set Rec1 = openDB.OpenRecord("SELECT StrID AS ID,StrName AS Name FROM Stream")
 
    rec.MoveFirst
    Set dtgCode.DataSource = rec
    dtgCode.Caption = "Details of A/L Streams"
    dtgCode.HeadFont.Bold = True
    dtgCode.HeadFont.Size = 10
    dtgCode.Columns(0).Alignment = dbgCenter
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Text1.SetFocus
rec.Close
Rec1.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub
