VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStaffSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   ControlBox      =   0   'False
   Icon            =   "frmStaffSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin MSComctlLib.ListView lvstaff 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16776960
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "StaffID"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label1 
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
      Left            =   840
      TabIndex        =   2
      Top             =   5040
      Width           =   3015
   End
End
Attribute VB_Name = "frmStaffSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStaff As ADODB.Recordset


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "Find Staff")

Set RecStaff = openDB.OpenRecord("select * from STAFF")


If Not (RecStaff.EOF And RecStaff.BOF) Then
    RecStaff.MoveFirst
        lvstaff.ListItems.clear
        While Not RecStaff.EOF
            Set List = lvstaff.ListItems.Add
            List.Text = RecStaff.Fields(0)
            List.SubItems(1) = RecStaff.Fields(1)
            RecStaff.MoveNext
        Wend
        Else
        lvstaff.ListItems.clear
End If
main.Enabled = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecStaff.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True

End Sub

Private Sub lvstaff_Click()
On Error Resume Next
Text1.Text = lvstaff.SelectedItem.Text + "   " + lvstaff.SelectedItem.SubItems(1)
If (Err.Number = 91) Then
MsgBox "No record found"
End If

End Sub

Private Sub lvstaff_DblClick()
On Error Resume Next
modform.Staff1 = lvstaff.SelectedItem.Text
If (modform.formname = "StaffLeave") Then
frmstaffleaves.dcStaffID = lvstaff.SelectedItem.Text
ElseIf (modform.formname = "StaffSch") Then
frmstaffScheduling.dcStaffID = lvstaff.SelectedItem.Text
End If
If (Err.Number = 91) Then
MsgBox "No record found"
Exit Sub
End If
Unload Me
End Sub

Private Sub lvstaff_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call lvstaff_DblClick
End If

End Sub
