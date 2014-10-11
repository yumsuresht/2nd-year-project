VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIndAvg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8130
   Icon            =   "frmIndAvg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8130
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7935
      Begin MSComctlLib.ListView lvstudent 
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "DblClick  to show report"
         Top             =   240
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16776960
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student Id"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Student Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Father Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date of Admission"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Admission Grade"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Press Enter or DblClick  to Select"
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
      TabIndex        =   6
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Student Name"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmIndAvg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStu As ADODB.Recordset
Dim s As String

Private Sub Command1_Click()
Call modReports.IndYearAvg(s)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "Student")

Set RecStu = openDB.OpenRecord("select M.StuID,M.StudentName,M.FatherName,M.D_Of_Admin,M.AdminGrade from MAINSTUDENTS M")


If Not (RecStu.EOF And RecStu.BOF) Then
    RecStu.MoveFirst
        lvstudent.ListItems.clear
        While Not RecStu.EOF
            Set List = lvstudent.ListItems.Add
            List.Text = RecStu.Fields(0)
            List.SubItems(1) = RecStu.Fields(1)
            List.SubItems(2) = RecStu.Fields(2)
            List.SubItems(3) = RecStu.Fields(3)
            List.SubItems(4) = RecStu.Fields(4)
            RecStu.MoveNext
        Wend
        Else
        lvstudent.ListItems.clear
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecStu.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub lvstudent_Click()
On Error Resume Next
Text1.Text = lvstudent.SelectedItem.SubItems(2) + " " + lvstudent.SelectedItem.SubItems(1)
s = lvstudent.SelectedItem.Text
If (Err.Number = 91) Then
MsgBox "No record found"
End If
End Sub

Private Sub lvstudent_DblClick()
Call Command1_Click
End Sub

Private Sub lvstudent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
