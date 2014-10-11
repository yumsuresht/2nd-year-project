VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmclub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clubs"
   ClientHeight    =   5145
   ClientLeft      =   3045
   ClientTop       =   1935
   ClientWidth     =   7725
   Icon            =   "frmclub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7725
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add"
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Edit"
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<<"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   ">>"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Frame Frame3 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   7575
      Begin MSDataGridLib.DataGrid dtgCode 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4683
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
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   529
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
      Format          =   76021763
      UpDown          =   -1  'True
      CurrentDate     =   38213
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Club ID"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   15
   End
   Begin VB.Label Label2 
      Caption         =   "Club Name"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Start Year"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Description"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmclub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset

Private Sub Command1_Click()
On Error Resume Next
If Not (Text2.Text = "") Then
    rec.MoveFirst
    rec.Find "Name = '" & Trim(Text2.Text) & "'"
    If rec.EOF Then
        rec.AddNew
        rec!name = UCase(Text2.Text)
        rec!StartYear = dtp.Value
        rec!DESCRIPTIONS = UCase(Text4.Text)
        rec.UpdateBatch
        dtgCode.Refresh
        main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
    Else
        MsgBox "Club already added....!"
    End If
        Text2.Text = ""
Else
    MsgBox "BLANK FIELD"
End If

End Sub

Private Sub Command2_Click()
On Error Resume Next

    If MsgBox("Do you really want to delete this record?", vbYesNo, "Delete") = vbYes Then
        rec.Delete
        rec.UpdateBatch
        rec.Requery
        dtgCode.Refresh
    End If
End Sub

Private Sub Command3_Click()
On Error GoTo Errhand
    rec!name = Text2.Text
    rec!StartYear = dtp.Value
    rec!DESCRIPTIONS = Text4.Text
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
Dim s As String
On Error Resume Next
s = InputBox("Enter Club Name", SEARCH)
If (s <> "") Then
rec.MoveFirst
    rec.Find "Name = '" & UCase(Trim(s)) & "'"
    If rec.EOF Then
        MsgBox "CAN NOT FIND"
    Else
        Text2.Text = rec!name
        dtp.Value = rec!StartYear
        Text4.Text = rec!DESCRIPTIONS
    End If
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
rec.MovePrevious
    If rec.BOF Then rec.MoveFirst
        Text2.Text = rec!name
        dtp.Value = rec!StartYear
        Text4.Text = rec!DESCRIPTIONS
End Sub

Private Sub Command7_Click()
On Error Resume Next
rec.MoveNext
    If rec.EOF Then rec.MoveLast
        Text2.Text = rec!name
        dtp.Value = rec!StartYear
        Text4.Text = rec!DESCRIPTIONS
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "ADD Club/Unions")
    Set rec = openDB.OpenRecord("SELECT CName AS Name,StartYear,Descriptions FROM CLUB ORDER BY CName")
    rec.MoveFirst
    Set dtgCode.DataSource = rec
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Text2.SetFocus
rec.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

