VERSION 5.00
Begin VB.Form frminitial 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6135
   Icon            =   "initialize.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton Command9 
         Caption         =   "Assign"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   24
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Assign"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   23
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Assign"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Assign"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Assign"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Assign"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Assign"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Temp ID"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Library Transaction No"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Book ReserveNo"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Library Receipt No"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Library MemberID"
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Staff ID"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Student ID"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "You can enter value only one time by this window, So please fill carefully"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   3735
      End
      Begin VB.Label Label9 
         Caption         =   "Warning..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frminitial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset

Private Sub Command1_Click()

End Sub



Private Sub Command10_Click()

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If (Text1.Text = "") Then
MsgBox "You must enter number value", vbInformation
Exit Sub
End If
rec!TEMIDS = Text1.Text
rec.UpdateBatch
rec.Requery
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
On Error Resume Next
If (Text3.Text = "") Then
MsgBox "You must enter number value", vbInformation
Exit Sub
End If
rec!STAFFIDS = Val(Text3.Text)
rec.UpdateBatch
rec.Requery
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
On Error Resume Next
If (Text5.Text = "") Then
MsgBox "You must enter number value", vbInformation
Exit Sub
End If

rec!RECEPID = Val(Text5.Text)
rec.UpdateBatch
rec.Requery
Command5.Enabled = False
End Sub

Private Sub Command6_Click()
On Error Resume Next
If (Text7.Text = "") Then
MsgBox "You must enter number value", vbInformation
Exit Sub
End If

rec!LIBTRANSNO = Text7.Text
rec.UpdateBatch
rec.Requery
Command6.Enabled = False
End Sub

Private Sub Command7_Click()
On Error Resume Next
If (Text2.Text = "") Then
MsgBox "You must enter number value", vbInformation
Exit Sub
End If

rec!STUIDS = Text2.Text
rec.UpdateBatch
rec.Requery
Command7.Enabled = False
End Sub

Private Sub Command8_Click()
On Error Resume Next
If (Text4.Text = "") Then
MsgBox "You must enter number value", vbInformation
Exit Sub
End If

rec!MEMBERIDS = Val(Text4.Text)
rec.UpdateBatch
rec.Requery
Command8.Enabled = False
End Sub

Private Sub Command9_Click()
On Error Resume Next
If (Text6.Text = "") Then
MsgBox "You must enter number value", vbInformation
Exit Sub
End If

rec!RESERVENO = Val(Text6.Text)
rec.UpdateBatch
rec.Requery
Command9.Enabled = False
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "Application Setup")
    Set rec = openDB.OpenRecord("SELECT * FROM IDS")
    rec.MoveLast
    If (rec!IDS = "") Then
    rec.AddNew
    rec!IDS = "1"
    rec.UpdateBatch
    rec.Requery
    End If
    
    rec.MoveLast
    
    Text1.Text = rec!TEMIDS
    Text2.Text = rec!STUIDS
    Text3.Text = rec!STAFFIDS
    Text4.Text = rec!MEMBERIDS
    Text5.Text = rec!RECEPID
    Text6.Text = rec!RESERVENO
    Text7.Text = rec!LIBTRANSNO
    
    If (Text1.Text = "") Then
    Command3.Enabled = True
    Text1.Locked = False
    End If
    
    If (Text2.Text = "") Then
    Command7.Enabled = True
    Text2.Locked = False
    End If
    
    If (Text3.Text = "") Then
    Command4.Enabled = True
    Text3.Locked = False
    End If
    
    If (Text4.Text = "") Then
    Command8.Enabled = True
    Text4.Locked = False
    End If
    
    If (Text5.Text = "") Then
    Command5.Enabled = True
    Text5.Locked = False
    End If
    
    If (Text6.Text = "") Then
    Command9.Enabled = True
    Text6.Locked = False
    End If
    
    If (Text7.Text = "") Then
    Command6.Enabled = True
    Text7.Locked = False
    End If
    
End Sub

Public Sub Buttoncontrol()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
rec.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Command3.SetFocus
Else
    msg = MsgBox("You must enter number value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Command7.SetFocus
Else
    msg = MsgBox("You must enter number value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Command4.SetFocus
Else
    msg = MsgBox("You must enter number value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Command8.SetFocus
Else
    msg = MsgBox("You must enter number value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Command5.SetFocus
Else
    msg = MsgBox("You must enter number value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Command9.SetFocus
Else
    msg = MsgBox("You must enter number value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Command6.SetFocus
Else
    msg = MsgBox("You must enter number value", vbExclamation)
    KeyAscii = 0
End If
End Sub
