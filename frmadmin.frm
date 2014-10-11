VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmadmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   Icon            =   "frmadmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "User"
      TabPicture(0)   =   "frmadmin.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Groups"
      TabPicture(1)   =   "frmadmin.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dcgroup"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Change Admin Password"
      TabPicture(2)   =   "frmadmin.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Command12"
      Tab(2).Control(2)=   "Command10"
      Tab(2).Control(3)=   "txtverpass"
      Tab(2).Control(4)=   "txtnewpass"
      Tab(2).Control(5)=   "txtoldpass"
      Tab(2).Control(6)=   "Label10"
      Tab(2).Control(7)=   "Label9"
      Tab(2).Control(8)=   "Label8"
      Tab(2).Control(9)=   "Label7"
      Tab(2).Control(10)=   "Label6"
      Tab(2).ControlCount=   11
      Begin VB.CommandButton Command9 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -69600
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   -74760
         TabIndex        =   29
         Top             =   2520
         Width           =   3495
         Begin VB.Label Label2 
            Caption         =   "Current User :  Admin"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Change"
         Height          =   375
         Left            =   -70200
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   -70200
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtverpass 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -73560
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtnewpass 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -73560
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtoldpass 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -73560
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   960
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo dcgroup 
         Height          =   315
         Left            =   -74040
         TabIndex        =   6
         Top             =   600
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Frame Frame3 
         Caption         =   "Permission"
         Height          =   3495
         Left            =   -74880
         TabIndex        =   22
         Top             =   1320
         Width           =   6375
         Begin VB.CheckBox Check12 
            Caption         =   "Old Student"
            Height          =   255
            Left            =   2640
            TabIndex        =   43
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Result"
            Height          =   255
            Left            =   2640
            TabIndex        =   42
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Student Leaving"
            Height          =   255
            Left            =   2640
            TabIndex        =   41
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CheckBox Check9 
            Caption         =   "New Student"
            Height          =   255
            Left            =   2640
            TabIndex        =   40
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Report"
            Height          =   255
            Left            =   5160
            TabIndex        =   39
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Library"
            Height          =   255
            Left            =   4080
            TabIndex        =   38
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Term Average"
            Height          =   255
            Left            =   2640
            TabIndex        =   37
            Top             =   2040
            Width           =   1455
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Student"
            Height          =   255
            Left            =   2640
            TabIndex        =   36
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Club Maintains"
            Height          =   255
            Left            =   360
            TabIndex        =   35
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Staff"
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   1440
            Width           =   975
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Master Setup"
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Administration"
            Height          =   255
            Left            =   360
            TabIndex        =   32
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   5160
            TabIndex        =   8
            Top             =   3000
            Width           =   1095
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Apply"
            Height          =   375
            Left            =   3840
            TabIndex        =   7
            Top             =   3000
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Group Membership"
         Height          =   3135
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   6375
         Begin VB.CommandButton Command6 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   5040
            TabIndex        =   5
            Top             =   2640
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Apply"
            Height          =   375
            Left            =   3840
            TabIndex        =   4
            ToolTipText     =   "Save & Exit"
            Top             =   2640
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Remove"
            Height          =   375
            Left            =   2520
            TabIndex        =   3
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Add"
            Height          =   375
            Left            =   2520
            TabIndex        =   2
            Top             =   840
            Width           =   1215
         End
         Begin VB.ListBox listmem 
            Appearance      =   0  'Flat
            Height          =   1785
            Left            =   3900
            TabIndex        =   20
            Top             =   600
            Width           =   2200
         End
         Begin VB.ListBox listgroup 
            Appearance      =   0  'Flat
            Height          =   1785
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   2200
         End
         Begin VB.Label Label4 
            Caption         =   "Member Of:"
            Height          =   255
            Left            =   3900
            TabIndex        =   19
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Available Groups:"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "User"
         Height          =   1215
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   6375
         Begin VB.CommandButton Command1 
            Caption         =   "Delete User"
            Height          =   375
            Left            =   4680
            TabIndex        =   44
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Clear Password"
            Height          =   375
            Left            =   4680
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dcid 
            Height          =   315
            Left            =   1200
            TabIndex        =   1
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label1 
            Caption         =   "UserID"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Admin"
         Height          =   255
         Left            =   -73560
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Verify :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "New Password :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   25
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Old Password :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "UserID :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Recgroup As ADODB.Recordset
Dim RecUser As ADODB.Recordset
Dim RecUser1 As ADODB.Recordset
Dim RecStaff As ADODB.Recordset
Dim memgro As String



Private Sub Command1_Click()
On Error Resume Next
RecUser1.Delete
RecUser1.UpdateBatch
RecUser1.Requery
dcid.Text = ""
If (Err.Number = 0) Then
MsgBox "Password is Successfully deleted"
End If
End Sub

Private Sub Command10_Click()
Unload Me
End Sub

Private Sub Command11_Click()
On Error Resume Next
RecUser1!Passwords = ""
RecUser1.UpdateBatch
RecUser1.Requery

If (Err.Number = 0) Then
MsgBox "Password is Successfully Cleared"
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
RecUser1.Find "StaffID = 'Admin'"
'If (txtoldpass.Text = "") Then
'MsgBox "You must enter the old password"
'Exit Sub
'End If
If (Trim(RecUser1!Passwords) <> Trim(txtoldpass.Text)) Then
    MsgBox "Your old password is incorrect"
    Call modform.ClearTextBoxes(Me)
    Exit Sub
Else
    If (txtnewpass.Text <> txtverpass.Text) Then
    'Or (txtnewpass.Text = "") Or (txtverpass.Text = "") Then
        MsgBox "Verify the new password by retyping it in the Verify box and clicking Change.", vbInformation
        Call modform.ClearTextBoxes(Me)

    Exit Sub
    Else
        RecUser1!Passwords = txtnewpass.Text
        RecUser1.UpdateBatch
        RecUser1.Requery
    End If
End If
Call modform.ClearTextBoxes(Me)

If (Err.Number = 0) Then
MsgBox "Admin Password Changed"
End If

End Sub

Private Sub Command13_Click()
MsgBox lstmnu.ListCount
For i = 1 To lstmnu.ListCount
Text1.Text = lstmnu.Text
Next i

End Sub

Private Sub Command2_Click()
On Error Resume Next
    If (Trim(dcid.Text) = "Admin") Then
        RecUser1.MoveFirst
        RecUser1.AddNew
        RecUser1!staffid = "Admin"
        RecUser1!Passwords = "Admin"
        RecUser1!GroupName = "Admin"
        RecUser1.UpdateBatch
        RecUser1.Requery
    ElseIf (listmem.ListCount = 0) Then
        MsgBox "You must select the group"
        Exit Sub
    ElseIf RecUser1.EOF Then
        RecUser1.MoveFirst
        RecUser1.AddNew
        RecUser1!staffid = Trim(dcid.Text)
        RecUser1!Passwords = ""
        RecUser1!GroupName = Trim(memgro)
        RecUser1.UpdateBatch
        RecUser1.Requery
    Else
    RecUser1!GroupName = Trim(memgro)
        RecUser1.UpdateBatch
        RecUser1.Requery
    
    End If
    
    If Not ((Err.Number = 0 Or Err.Number = 3021)) Then
    MsgBox "already added"
    Exit Sub
    Else
    Unload Me
    End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim groname As String
memgro = listgroup.Text
If (listmem.ListCount = 0) Then
    If (MsgBox("Are you sure to select this group?", vbOKCancel) = vbOK) Then
        groname = listgroup.Text
        If listgroup.ListIndex >= 0 Then
            listmem.AddItem listgroup.Text
        End If
    Else
        Exit Sub
    End If
Else
        MsgBox "This user already registered"
Exit Sub
End If



End Sub

Private Sub Command4_Click()

Dim groname As String

If listmem.SelCount = 1 Then
        listmem.RemoveItem listmem.ListIndex
ElseIf listmem.SelCount > 1 Then
    For i = listmem.ListCount - 1 To 0 Step -1
        If listmem.Selected(i) Then
        listmem.RemoveItem i
        End If
    Next
End If
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
On Error Resume Next
If (Trim(dcgroup.Text) = "") Then
MsgBox "Check your User group"
Exit Sub
End If

Recgroup.Find "GroupName = '" & Trim(dcgroup.Text) & "'"
If Recgroup.EOF Then
    Recgroup.AddNew
    Recgroup!GroupName = Trim(dcgroup.Text)
    
    Recgroup!Administrator = Check1.Value
    Recgroup!MasterSetup = Check2.Value
    Recgroup!Staff = Check3.Value
    Recgroup!ClubMaintain = Check4.Value
    Recgroup!Students = Check5.Value
    Recgroup!NewStudent = Check9.Value
    Recgroup!StudentLea = Check10.Value
    Recgroup!Result = Check11.Value
    Recgroup!TermAve = Check6.Value
    Recgroup!OldStudnt = Check12.Value
    Recgroup!Library = Check7.Value
    Recgroup!Report = Check8.Value
    
    Recgroup.UpdateBatch
    Recgroup.Requery
Else
    Recgroup!Administrator = Check1.Value
    Recgroup!MasterSetup = Check2.Value
    Recgroup!Staff = Check3.Value
    Recgroup!ClubMaintain = Check4.Value
    Recgroup!Students = Check5.Value
    Recgroup!NewStudent = Check9.Value
    Recgroup!StudentLea = Check10.Value
    Recgroup!Result = Check11.Value
    Recgroup!TermAve = Check6.Value
    Recgroup!OldStudnt = Check12.Value
    Recgroup!Library = Check7.Value
    Recgroup!Report = Check8.Value
    Recgroup.UpdateBatch
    Recgroup.Requery
End If

    Call Adddata(listgroup, Recgroup)


If ((Err.Number = 0 Or Err.Number = 3021)) Then
Call Adddata(listgroup, Recgroup)
    MsgBox "User Group Sucessfully Added"
End If
End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub


Private Sub Command9_Click()
On Error Resume Next
If (dcgroup.Text = "Admin") Then
MsgBox "You can not delete Admin Group"
Exit Sub
End If


If (dcgroup.Text = "") Then
MsgBox "Select the Group and click Delete"
Else
Recgroup.Delete
Recgroup.UpdateBatch
Recgroup.Requery
dcgroup.Text = ""

End If
Recgroup.Requery
If (Err.Number = 0) Then
Call Adddata(listgroup, Recgroup)
MsgBox "User Group is Sucessfully Deleted"
Unload Me
End If
End Sub

Private Sub dcgroup_Change()
On Error Resume Next
Dim adminno As String
adminno = Trim(dcgroup.Text)
    If adminno <> "" Then
    Recgroup.MoveFirst
    Recgroup.Find "GroupName = '" & Trim(dcgroup.Text) & "'"
       If Recgroup.EOF Then
            Command9.Enabled = False
       Else
       Text1.Text = Recgroup!Administrator
            Call fillcheck
            Command9.Enabled = True

       End If
     End If


End Sub

Private Sub dcid_Change()
On Error Resume Next
RecUser1.MoveFirst
RecUser.MoveFirst
stfid = Trim(dcid.Text)
Set RecUser = openDB.OpenRecord("select * from LOGIN where StaffID='" & dcid.Text & "'")
Call Adddata(listmem, RecUser)
RecUser1.Find "StaffID = '" & stfid & "'"
End Sub

Private Sub dcuserid_Change()

End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "User and Group Accounts")
main.Enabled = False

Set Recgroup = openDB.OpenRecord("select * from USERGROUP")
Set RecUser1 = openDB.OpenRecord("select * from LOGIN")
Set RecStaff = openDB.OpenRecord("select * from STAFF")


Recgroup.MoveFirst
dcgroup.ListField = "GroupName"
Set dcgroup.RowSource = Recgroup
    
RecUser1.MoveFirst
dcid.ListField = "StaffID"
Set dcid.RowSource = RecUser1



Call Adddata(listgroup, Recgroup)
Command9.Enabled = False
Command13.Enabled = True
Call AddMnu



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Recgroup.Close
RecUser.Close
RecUser1.Close
RecStaff.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True
End Sub


Public Sub Adddata(groupna As ListBox, rec As ADODB.Recordset)
groupna.clear
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        While Not rec.EOF
            groupna.AddItem rec!GroupName
            rec.MoveNext
        Wend
    End If
End Sub

Private Sub List1_Click()

End Sub
Public Sub AddMnu()
    lstmnu.AddItem "File"
    lstmnu.AddItem "School"
    lstmnu.AddItem "Student"
    lstmnu.AddItem "Library"
            
End Sub

Public Sub Adddata1(name As String)
If (name = "Library") Then
    Recgroup!Library = 1
    Recgroup.UpdateBatch
End If
End Sub


Public Sub fillcheck()
On Error Resume Next
 Check1.Value = reply(Recgroup!Administrator)
 Check2.Value = reply(Recgroup!MasterSetup)
 Check3.Value = reply(Recgroup!Staff)
 Check4.Value = reply(Recgroup!ClubMaintain)
 Check5.Value = reply(Recgroup!Students)
 Check6.Value = reply(Recgroup!TermAve)
 Check7.Value = reply(Recgroup!Library)
 Check8.Value = reply(Recgroup!Report)
 Check9.Value = reply(Recgroup!NewStudent)
 Check10.Value = reply(Recgroup!StudentLea)
 Check11.Value = reply(Recgroup!Result)
 Check6.Value = reply(Recgroup!TermAve)
 Check12.Value = reply(Recgroup!OldStudnt)
 
 
 
End Sub


Public Function reply(aa As Boolean) As Integer
If (aa = True) Then
reply = 1
Else
reply = 0
End If

End Function

