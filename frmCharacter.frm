VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCharacter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   Icon            =   "frmCharacter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6165
   Begin MSComctlLib.ListView lvstu 
      Height          =   3255
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Press Enter or DblClick  to show the certificates"
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5741
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "StuID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FatherName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Street"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "City"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Admin Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Leave Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show Character Certificate"
      Height          =   855
      Left            =   4440
      Picture         =   "frmCharacter.frx":030A
      TabIndex        =   9
      ToolTipText     =   "Preview"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4215
      Begin VB.OptionButton optName 
         Caption         =   "Name"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optFaName 
         Caption         =   "Father Name"
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optStuId 
         Caption         =   "StuID"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "Address"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Applicant's"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Find What"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStu As ADODB.Recordset
Public STUID1 As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim condition As String

If (optAll.Value = True) Then
condition = ""
ElseIf (Text1.Text = "") Then
MsgBox "You must enter the search value"
Exit Sub
Else
    If (optStuId.Value = True) Then
        condition = "and M.StuID='" & Trim(Text1.Text) & "'"
    ElseIf (optAdd.Value = True) Then
        condition = "and M.Street like '%" & Trim(Text1.Text) & "%' or M.City like '%" & Trim(Text1.Text) & "%'"
    ElseIf (optName.Value = True) Then
        condition = "and M.StudentName like '%" & Trim(Text1.Text) & "%'"
    ElseIf (optFaName.Value = True) Then
        condition = "and M.FatherName like '%" & Trim(Text1.Text) & "%'"
    End If

End If
'Set RecStu = openDB.OpenRecord("select distinct(M.StuID) AS Student_ID,M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin AS Admission_Date,O.D_Of_Leave AS Leave_Date from MAINSTUDENTS M, OLDBOYS O where M.StuID=O.StuID " & condition)
'Set RecStu = openDB.OpenRecord("select distinct(M.StuID) AS Student_ID,M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin AS Admission_Date,O.D_Of_Leave AS Leave_Date from MAINSTUDENTS M, OLDBOYS O where M.StuID=O.StuID " & condition)

Set RecStu = openDB.OpenRecord("select distinct(M.StuID),M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin AS Admission_Date from MAINSTUDENTS M, OLDBOYS O where M.StuID=O.StuID " & condition & " Union select distinct(M.StuID),M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin AS Admission_Date from MAINSTUDENTS M, ACTIVESTUDENT A where M.StuID=A.StuID " & condition)

If (RecStu.RecordCount = 0 Or Err.Number <> 0) Then
MsgBox "Can not find"
Exit Sub
End If

Call filllist(RecStu)

End Sub

Private Sub Command3_Click()

If (Text1.Text <> "") Then
 STUID1 = Text1.Text
    If (Command3.Caption = "Show Character Certificate") Then
        frmcertificate.Show
    End If
    If (Command3.Caption = "O/L Result Sheet") Then
        OlResultSheet (STUID1)
    End If
    
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "Student Search")


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecStu.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Public Sub filllist(RecStu1 As ADODB.Recordset)
On Error Resume Next
If Not (RecStu1.EOF And RecStu1.BOF) Then
    RecStu1.MoveFirst
        lvstu.ListItems.clear
        While Not RecStu1.EOF
            Set List = lvstu.ListItems.Add
            List.Text = RecStu1.Fields(0)
            List.SubItems(1) = RecStu1.Fields(1)
            List.SubItems(2) = RecStu1.Fields(2)
            List.SubItems(3) = RecStu1.Fields(3)
            List.SubItems(4) = RecStu1.Fields(4)
            List.SubItems(5) = RecStu1.Fields(5)
            List.SubItems(6) = RecStu1.Fields(6)
            RecStu1.MoveNext
        Wend
        Else
        lvstu.ListItems.clear
End If
End Sub

Private Sub lvstu_Click()
On Error Resume Next
Text1.Text = lvstu.SelectedItem.Text
End Sub

Private Sub lvstu_DblClick()
On Error Resume Next
Text1.Text = lvstu.SelectedItem.Text
Call Command3_Click
End Sub

Private Sub lvstu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call lvstu_DblClick
End If
End Sub
