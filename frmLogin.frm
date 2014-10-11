VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Text            =   "Admin"
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Dim Reclogin As ADODB.Recordset

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    main.Enabled = True
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
If (txtUserName.Text = "Admin") Then
    Reclogin.MoveFirst
    Reclogin.Find "StaffID = 'Admin'"
    If Reclogin.EOF Or Reclogin!Passwords <> Trim(txtPassword.Text) Then
        MsgBox "Invalid Administrator Password", vbInformation, "Login Fail"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        LoginSucceeded = False
    ElseIf (Reclogin!Passwords = Trim(txtPassword.Text)) Then
        LoginSucceeded = True
        main.Enabled = True
        Call frmtermavg.viewterms
        frmtermavg.dgterm.Refresh
        Unload Me
    End If
Else
    LoginSucceeded = False
    MsgBox "Only the Administrator can change the previous Term Marks", vbInformation, "Login"
End If
End Sub

Private Sub Form_Load()
Set Reclogin = openDB.OpenRecord("SELECT * FROM LOGIN")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Reclogin.Close
End Sub
