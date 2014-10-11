VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStaffSearch1 
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   Icon            =   "frmStaffSearch1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtcontent 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   3375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Search"
         Height          =   1695
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   4335
         Begin VB.OptionButton optFaName 
            Caption         =   "Religion"
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton optName 
            Caption         =   "Name"
            Height          =   255
            Left            =   720
            TabIndex        =   8
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton optAdd 
            Caption         =   "City"
            Height          =   255
            Left            =   3000
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optTemId 
            Caption         =   "Staff ID"
            Height          =   255
            Left            =   1560
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optAll 
            Caption         =   "All"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.Frame Frame2 
            Caption         =   "Applicant's"
            Height          =   855
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   4095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Operations"
         Height          =   2295
         Left            =   4560
         TabIndex        =   1
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   840
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView listteasub 
         Height          =   2895
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   6135
         _ExtentX        =   10821
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
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "StaffID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "City"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Grade"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Salary"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Working Hours"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Religion"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Find What :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmStaffSearch1"
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
        condition = "where StaffID='" & Trim(txtcontent.Text) & "'"
    ElseIf (optAdd.Value = True) Then
        condition = "where City like '%" & Trim(txtcontent.Text) & "%'"
    ElseIf (optName.Value = True) Then
        condition = "where FullName like '%" & Trim(txtcontent.Text) & "%'"
    ElseIf (optFaName.Value = True) Then
       condition = "where Religion like '%" & Trim(txtcontent.Text) & "%'"
    End If

End If
Set RecTemp = openDB.OpenRecord("select * from STAFF " & condition)
If (RecTemp.RecordCount = 0 Or Err.Number <> 0) Then
MsgBox "Cannot find"
Exit Sub
End If
Call filllist


Set dgTemp.DataSource = RecTemp
   

End Sub

Private Sub Form_Load()
On Error Resume Next
main.Enabled = False

Call modform.FormSize(Me, "Staff Search")
    

    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecTemp.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True

End Sub
Public Sub filllist()
If Not (RecTemp.EOF And RecTemp.BOF) Then
    RecTemp.MoveFirst
        listteasub.ListItems.clear
        While Not RecTemp.EOF
            Set List = listteasub.ListItems.Add
            List.Text = RecTemp.Fields(0)
            List.SubItems(1) = RecTemp.Fields(1)
            List.SubItems(2) = RecTemp!City
            List.SubItems(3) = RecTemp!grade
            List.SubItems(4) = RecTemp!Salary
            List.SubItems(5) = RecTemp!Work_Hours
            List.SubItems(6) = RecTemp!Religion
            RecTemp.MoveNext
        Wend
        Else
        listteasub.ListItems.clear
    End If
End Sub

Private Sub listteasub_DblClick()
modform.Staff1 = listteasub.SelectedItem.Text
Call frmnewstaff.filllists
Unload Me


End Sub

Private Sub txtcontent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmdFind_Click
End If

End Sub
