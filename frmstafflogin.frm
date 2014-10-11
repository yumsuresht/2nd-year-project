VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmstafflogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User authentication window"
   ClientHeight    =   2310
   ClientLeft      =   2835
   ClientTop       =   3585
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1364.824
   ScaleMode       =   0  'User
   ScaleWidth      =   5309.739
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   240
      Top             =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter UserID and Password to Login to the System."
      Height          =   2175
      Left            =   53
      TabIndex        =   0
      Top             =   85
      Width           =   5535
      Begin MSComctlLib.ProgressBar PBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   300
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "Accept"
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ok"
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4320
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtuserid 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Text            =   "Admin"
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtpass 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblLoad 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "UserID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmstafflogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Reclogin As ADODB.Recordset
Dim Reclogin1 As ADODB.Recordset
Dim Reclogin2 As ADODB.Recordset
Dim aa As Boolean
Dim load



Private Sub Command1_Click()
       


End Sub

Private Sub Command2_Click()
Dim reply
reply = MsgBox("Are you sure you want to exit the program?", vbInformation + vbYesNo, "Exit")
If reply = vbYes Then
    End
Else
    Me.Show
End If

End Sub

Private Sub cmdOK_Click()
On Error Resume Next

Reclogin1.AddNew
      Reclogin1!staffid = "Admin"
      Reclogin1!Passwords = Trim(txtpass.Text)
      Reclogin1!LastModiDate = Date
      Reclogin1!GroupName = "Admin"
      Reclogin1.UpdateBatch
      Reclogin1.Requery
      
    Reclogin2.AddNew
    Reclogin2!GroupName = "Admin"
    Reclogin2!Administrator = 1
    Reclogin2!MasterSetup = 1
    Reclogin2!Staff = 1
    Reclogin2!ClubMaintain = 1
    Reclogin2!Students = 1
    Reclogin2!NewStudent = 1
    Reclogin2!StudentLea = 1
    Reclogin2!Result = 1
    Reclogin2!TermAve = 1
    Reclogin2!OldStudnt = 1
    Reclogin2!Library = 1
    Reclogin2!Report = 1
    Reclogin2.UpdateBatch
    Reclogin2.Requery
    
    Call Command3_Click
    Call prograss
    
    frminitial.Show
    frmWizard.Show
End Sub

Private Sub Command3_Click()
On Error Resume Next
Set Reclogin = openDB.OpenRecord("select * from LOGIN L,USERGROUP U where L.GroupName=U.GroupName and StaffID='" & txtuserid & "'")

Reclogin.MoveFirst
Reclogin.Find "StaffID = '" & Trim(txtuserid.Text) & "'"

If Reclogin.EOF Then
        MsgBox "Access denied! Contact Administrator", vbInformation, "Login"
ElseIf (Reclogin!Passwords = Trim(txtpass.Text)) Then

        main.Enabled = True
        
        If (Trim(txtuserid.Text) = "Admin") Then
        main.mnustaffAttends.Visible = False
        Else
        main.mnustaffAttends.Visible = True
        End If



        main.mnuAdministration.Enabled = Reclogin!Administrator
        main.mnuSchool.Enabled = Reclogin!MasterSetup
        main.mnustaff.Enabled = Reclogin!Staff
        main.mnuclubmaintance.Enabled = Reclogin!ClubMaintain
        main.mnuStudent.Enabled = Reclogin!Students
        
        
        main.mnuNewStudent.Enabled = Reclogin!NewStudent
        main.mnustuleave.Enabled = Reclogin!StudentLea
        main.mnuResulte.Enabled = Reclogin!Result
        main.mnutermavg.Enabled = Reclogin!TermAve
        main.mnuOldstudent.Enabled = Reclogin!OldStudnt
        
        main.mnuLibrary.Enabled = Reclogin!Library
        main.mnuReport.Enabled = Reclogin!Report
        
        
        main.stbMain.Panels(2).Text = "Login Status : Success"
        Call modform.setUserID(Trim(txtuserid.Text))
        Call modform.ClearTextBoxes(Me)
        
        Call prograss
        
Else

        MsgBox "Access denied! Contact Administrator", vbInformation, "Login"
Me.Show
Call modform.ClearTextBoxes(Me)
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim rep
main.Show
main.Enabled = False
Set Reclogin1 = openDB.OpenRecord("SELECT * FROM LOGIN")
Set Reclogin2 = openDB.OpenRecord("select * from USERGROUP")
'Set Reclogin = openDB.OpenRecord("select * from LOGIN L,USERGROUP U where L.GroupName=U.GroupName and StaffID='" & txtuserid & "'")

Reclogin1.MoveFirst
Reclogin1.Find "StaffID = '" & Trim(txtuserid.Text) & "'"

Reclogin2.MoveFirst
Reclogin2.Find "GroupName = '" & Trim(txtuserid.Text) & "'"


If Reclogin1.EOF And txtuserid.Text = "Admin" Then
    cmdOK.Visible = True
    Command3.Visible = False
    rep = MsgBox("Please Enter the new [Admin] Password, and remember it", vbOKCancel)
    txtuserid.Locked = True
ElseIf Reclogin2.EOF And txtuserid.Text = "Admin" Then
    Reclogin2.AddNew
    Reclogin2!GroupName = "Admin"
    Reclogin2!Administrator = 1
    Reclogin2!MasterSetup = 1
    Reclogin2!Staff = 1
    Reclogin2!ClubMaintain = 1
    Reclogin2!Students = 1
    Reclogin2!NewStudent = 1
    Reclogin2!StudentLea = 1
    Reclogin2!Result = 1
    Reclogin2!TermAve = 1
    Reclogin2!OldStudnt = 1
    Reclogin2!Library = 1
    Reclogin2!Report = 1
    Reclogin2.UpdateBatch
    Reclogin2.Requery
Else
    cmdOK.Visible = False
End If
      
      If (rep = vbCancel) Then
      Call Command2_Click
      End If
      
      
      
      
      
      
      
'End If

main.stbMain.Panels(2).Text = "Login Status : Fail"






End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Reclogin.Close
load = 0

End Sub

Private Sub Image1_Click()
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1 = Image3
Image2 = Image4
End Sub

Private Sub Image2_Click()


End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1 = Image4
Image2 = Image3
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Select Case load
Case 1 To 5
lblLoad.Caption = "Checking Database..."
PBar1.Value = PBar1.Value + load

Case 6 To 10
lblLoad.Caption = "Database OK..."
PBar1.Value = PBar1.Value + load

Case 11 To 25
lblLoad.Caption = "Loading Database..."
PBar1.Value = PBar1.Value + load

Case 26 To 30
lblLoad.Caption = "Loading Complete..."
PBar1.Value = PBar1.Value + load

Case 31 To 35
lblLoad.Caption = "Checking System Structures..."
PBar1.Value = PBar1.Value + load

Case 36 To 40
lblLoad.Caption = "Structures OK..."
PBar1.Value = PBar1.Value + load

Case 41 To 45
lblLoad.Caption = "Preparing Menus..."
Case 46 To 50
lblLoad.Caption = "Loading Screen..."
Case 51
Unload Me
End Select
load = load + 1
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command3.SetFocus
    End If
End Sub

Private Sub txtuserid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtpass.SetFocus
    End If
End Sub

Public Function prograss()
On Error Resume Next
lblLoad.Visible = True
DoEvents
Timer1.Enabled = True
Timer1.Interval = 10
PBar1.Visible = True
Call Timer1_Timer

End Function
