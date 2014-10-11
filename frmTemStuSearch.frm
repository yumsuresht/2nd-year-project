VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTemStuSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmTemStuSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6615
   Begin VB.Frame Frame4 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.Frame Frame3 
         Caption         =   "Operations"
         Height          =   2295
         Left            =   4560
         TabIndex        =   10
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Search"
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4335
         Begin VB.OptionButton optAll 
            Caption         =   "All"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optTemId 
            Caption         =   "TempID"
            Height          =   255
            Left            =   1560
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optAdd 
            Caption         =   "Address"
            Height          =   255
            Left            =   3000
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optName 
            Caption         =   "Name"
            Height          =   255
            Left            =   1320
            TabIndex        =   4
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton optFaName 
            Caption         =   "Father Name"
            Height          =   255
            Left            =   2640
            TabIndex        =   3
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Frame Frame2 
            Caption         =   "Applicant's"
            Height          =   855
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   4095
         End
      End
      Begin VB.TextBox txtcontent 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin MSComctlLib.ListView listteasub 
         Height          =   3375
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TemID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "StudentName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FatherName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "D_Of_Birth"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "City"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Find What :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTemStuSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecTemp As ADODB.Recordset



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
Dim condition As String

If (optAll.Value = True) Then
condition = ""
ElseIf (txtcontent.Text = "") Then
MsgBox "You must enter the search value"
Exit Sub
Else
    If (optTemId.Value = True) Then
        condition = "where TemID='" & Trim(txtcontent.Text) & "'"
    ElseIf (optAdd.Value = True) Then
        condition = "where City like '%" & Trim(txtcontent.Text) & "%'"
    ElseIf (optName.Value = True) Then
        condition = "where StudentName like '%" & Trim(txtcontent.Text) & "%'"
    ElseIf (optFaName.Value = True) Then
        condition = "where FatherName like '%" & Trim(txtcontent.Text) & "%'"
    End If

End If
Set RecTemp = openDB.OpenRecord("select * from TEMPSTUDENTS " & condition)
If (RecTemp.RecordCount = 0 Or Err.Number <> 0) Then
MsgBox "Can not find"
Exit Sub
End If



    
Call filllist
End Sub

Private Sub Form_Load()
On Error Resume Next
main.Enabled = False

Call modform.FormSize(Me, "Application Search")
    

    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecTemp.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Public Sub filllist()
If Not (RecTemp.EOF And RecTemp.BOF) Then
    RecTemp.MoveFirst
        listteasub.ListItems.clear
        While Not RecTemp.EOF
            Set List = listteasub.ListItems.Add
            List.Text = RecTemp.Fields(0)
            List.SubItems(1) = RecTemp.Fields(1)
            List.SubItems(2) = RecTemp.Fields(2)
            List.SubItems(3) = RecTemp.Fields(3)
            List.SubItems(4) = RecTemp!City
            RecTemp.MoveNext
        Wend
        Else
        listteasub.ListItems.clear
    End If
End Sub

Private Sub listteasub_DblClick()
On Error Resume Next
Call frmnewstu.disp(listteasub.SelectedItem.Text)
Call frmnewstu.disp1(listteasub.SelectedItem.Text)

End Sub
