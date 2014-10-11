VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAlresult 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6630
   Icon            =   "frmAlresult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6630
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   6375
      Begin VB.TextBox txtgen 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         MaxLength       =   1
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txteng 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         MaxLength       =   1
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtsub3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtsub2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtsub1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblgen 
         Caption         =   "GENERAL KNOWLEDGE"
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lbleng 
         Caption         =   "ENGLISH"
         Height          =   255
         Left            =   3600
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblsub3 
         Caption         =   "SUBJECT3"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblsub2 
         Caption         =   "SUBJECT2"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblsub1 
         Caption         =   "SUBJECT1"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.TextBox txtstr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtyear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin MSDataListLib.DataCombo dcadmin 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtindex 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Z-SCORE"
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "ISLAND RANK"
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "DISTRICT RANK"
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "STREAM"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "YEAR"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "INDEX NO"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ADMISSION NO"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAlresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecFollow As ADODB.Recordset
Dim RecAlRes As ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If (dcadmin.Text = "") Or (Trim(txtindex.Text) = "") Or (Trim(txtyear.Text) = "") Then
MsgBox "Please fill the particular fields"
Exit Sub
End If

If (RecAlRes.RecordCount = 2) Then
MsgBox "Can not insert this result, Student can take A/L by school maximum 2 times"
Call modform.ClearTextBoxes(Me)
Exit Sub
Else
RecAlRes.MoveFirst
RecAlRes.AddNew
RecAlRes!stuid = Trim(dcadmin.Text)
RecAlRes!StrName = Trim(txtstr.Text)
RecAlRes!IndexNo = Trim(txtindex.Text)
RecAlRes!Alyear = Trim(txtyear.Text)
RecAlRes!Subject1 = UCase(Trim(txtsub1.Text))
RecAlRes!Subject2 = UCase(Trim(txtsub2.Text))
RecAlRes!Subject3 = UCase(Trim(txtsub3.Text))
RecAlRes!English = UCase(Trim(txteng.Text))
RecAlRes!General = UCase(Trim(txtgen.Text))
RecAlRes!DistrictRank = Trim(Text1.Text)
RecAlRes!IslandRank = Trim(Text2.Text)
RecAlRes!ZScore = Val(Text3.Text)
RecAlRes.UpdateBatch
RecAlRes.Requery
Call modform.ClearTextBoxes(Me)
End If
If Not ((Err.Number = 0) Or (Err.Number = 3021)) Then
MsgBox "You already enter this details, Please check your details"
Exit Sub
End If
End Sub

Private Sub dcadmin_Change()
On Error Resume Next
Dim adminno As String
    adminno = Trim(dcadmin.Text)
    
               If adminno <> "" Then
                RecFollow.MoveFirst
                RecFollow.Find "StuID = '" & adminno & "'"
                If RecFollow.EOF Then
                    txtstr.Text = ""
                    lblsub1.Caption = "SUBJECT1"
                    lblsub2.Caption = "SUBJECT2"
                    lblsub3.Caption = "SUBJECT3"
                  
                Else
                    txtstr.Text = RecFollow!StrName
                    lblsub1.Caption = RecFollow!Subject1
                    lblsub2.Caption = RecFollow!Subject2
                    lblsub3.Caption = RecFollow!Subject3
                    Set RecAlRes = openDB.OpenRecord("select * from ALRESULT where StuId='" & adminno & "'")

               End If
            End If
 
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "A/L Result")
    Set RecFollow = openDB.OpenRecord("select * from FOLLOWSTREAM F, ACTIVESTUDENT A WHERE F.StuID=A.StuID")
    
    RecFollow.MoveFirst
    dcadmin.ListField = "StuID"
    Set dcadmin.RowSource = RecFollow
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecFollow.Close
RecAlRes.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Text2.SetFocus
Else
    msg = MsgBox("District Rank should be a numeric value", vbExclamation)
    KeyAscii = 0
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Text3.SetFocus
Else
    msg = MsgBox("Island Rank should be a numeric value", vbExclamation)
    KeyAscii = 0
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Command2.SetFocus
Else
    msg = MsgBox("Z-Score should be a numeric value", vbExclamation)
    KeyAscii = 0
End If

End Sub

Private Sub txteng_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
    msg = MsgBox("You can't enter this value", vbExclamation)
    KeyAscii = 0
ElseIf (KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 65 Or KeyAscii = 66 Or KeyAscii = 67 Or KeyAscii = 70 Or KeyAscii = 83 Or KeyAscii = 13 Or KeyAscii = 109 Or KeyAscii = 97 Or KeyAscii = 98 Or KeyAscii = 99 Or KeyAscii = 102 Or KeyAscii = 115) Then
    If (KeyAscii = 13) Then
    txtgen.SetFocus
    End If
Else
    msg = MsgBox("You can't enter this value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub txtindex_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    txtyear.SetFocus
Else
    msg = MsgBox("Index number should be a numeric value", vbExclamation)
    KeyAscii = 0
End If

End Sub


Private Sub txtsub1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
    msg = MsgBox("You can't enter this value", vbExclamation)
    KeyAscii = 0
ElseIf (KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 65 Or KeyAscii = 66 Or KeyAscii = 67 Or KeyAscii = 70 Or KeyAscii = 83 Or KeyAscii = 13 Or KeyAscii = 109 Or KeyAscii = 97 Or KeyAscii = 98 Or KeyAscii = 99 Or KeyAscii = 102 Or KeyAscii = 115) Then
    If (KeyAscii = 13) Then
    txtsub2.SetFocus
    End If
Else
    msg = MsgBox("You can't enter this value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub txtsub2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
    msg = MsgBox("You can't enter this value", vbExclamation)
    KeyAscii = 0
ElseIf (KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 65 Or KeyAscii = 66 Or KeyAscii = 67 Or KeyAscii = 70 Or KeyAscii = 83 Or KeyAscii = 13 Or KeyAscii = 109 Or KeyAscii = 97 Or KeyAscii = 98 Or KeyAscii = 99 Or KeyAscii = 102 Or KeyAscii = 115) Then
    If (KeyAscii = 13) Then
    txtsub3.SetFocus
    End If
Else
    msg = MsgBox("You can't enter this value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub txtsub3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
    msg = MsgBox("You can't enter this value", vbExclamation)
    KeyAscii = 0
ElseIf (KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 65 Or KeyAscii = 66 Or KeyAscii = 67 Or KeyAscii = 70 Or KeyAscii = 83 Or KeyAscii = 13 Or KeyAscii = 109 Or KeyAscii = 97 Or KeyAscii = 98 Or KeyAscii = 99 Or KeyAscii = 102 Or KeyAscii = 115) Then
    If (KeyAscii = 13) Then
    txteng.SetFocus
    End If
Else
    msg = MsgBox("You can't enter this value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub txtyear_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    txtsub1.SetFocus
Else
    msg = MsgBox("Year should be a numeric value ", vbExclamation)
    KeyAscii = 0
End If

End Sub
