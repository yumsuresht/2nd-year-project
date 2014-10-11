VERSION 5.00
Begin VB.Form frmchangepass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "frmchangepass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Verify :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "New Password :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Old Password :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "UserID :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmchangepass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Reclogin As ADODB.Recordset
Private Sub Command1_Click()
On Error Resume Next
If (Text2.Text <> Text3.Text) Then
'Or (Text2.Text = "") Or (Text3.Text = "") Then
        MsgBox "Verify the new password by retyping it in the Verify box and clicking Change.", vbInformation
        Call modform.ClearTextBoxes(Me)
Exit Sub
ElseIf (Reclogin!Passwords = Trim(Text1.Text)) Then
    Reclogin!Passwords = Text3.Text
    Reclogin.UpdateBatch
    Reclogin.Requery
Else
MsgBox "Old password is incorrect"
Exit Sub
End If

If (Err.Number = 0) Then
MsgBox "Password Changed"
Unload Me
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim useri As String
Call modform.FormSize(Me, "Change Password")
main.Enabled = False
useri = modform.userid
Label5.Caption = modform.userid
Set Reclogin = openDB.OpenRecord("select * from LOGIN where StaffID='" & Trim(Label5.Caption) & "'")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Reclogin.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True
End Sub

