VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOlResult 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   Icon            =   "frmOlResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6825
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   6615
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         TabIndex        =   29
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dcSID 
         Height          =   315
         Left            =   2040
         TabIndex        =   30
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "* Student ID"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Student Name"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "* Index No"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "* Year"
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Result"
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   6615
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   21
         Top             =   960
         Width           =   490
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   20
         Top             =   480
         Width           =   490
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   19
         Top             =   2880
         Width           =   490
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   18
         Top             =   2400
         Width           =   490
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1920
         Width           =   490
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1440
         Width           =   490
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   15
         Top             =   960
         Width           =   490
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1920
         Width           =   490
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   13
         Top             =   1440
         Width           =   490
      End
      Begin MSDataListLib.DataCombo dcLanguage 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         Text            =   "Language"
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   2
         Top             =   480
         Width           =   490
      End
      Begin MSDataListLib.DataCombo dcReligion 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         Text            =   "Religion"
      End
      Begin MSDataListLib.DataCombo dcAesthetic 
         Height          =   315
         Left            =   3600
         TabIndex        =   9
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         Text            =   "Aesthetic"
      End
      Begin MSDataListLib.DataCombo dcTechnical 
         Height          =   315
         Left            =   3600
         TabIndex        =   10
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         Text            =   "Technical"
      End
      Begin MSDataListLib.DataCombo dcAddtional1 
         Height          =   315
         Left            =   3600
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         Text            =   "Addtional1"
      End
      Begin MSDataListLib.DataCombo dcAddtional2 
         Height          =   315
         Left            =   3600
         TabIndex        =   12
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         Text            =   "Addtional2"
      End
      Begin VB.Label Label10 
         Caption         =   "*"
         Height          =   135
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "*"
         Height          =   135
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label23 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4560
         TabIndex        =   24
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label Label22 
         Caption         =   "Mandatory Fields"
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
         Left            =   4680
         TabIndex        =   23
         Top             =   3165
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "* English"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "* Social Studies"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "* Science and Technology"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "* Mathematics"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
   End
End
Attribute VB_Name = "frmOlResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecOlStu As ADODB.Recordset
Dim RecOlRes As ADODB.Recordset

Dim RecAddtionl As ADODB.Recordset
Dim RecAest As ADODB.Recordset
Dim RecLang As ADODB.Recordset
Dim RecReligions As ADODB.Recordset
Dim RecTech As ADODB.Recordset




Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next


If (Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "") Then
MsgBox "Please fill the mandatory fields"
Exit Sub
End If

If (RecOlRes.RecordCount = 2) Then
MsgBox "Can not insert this result, Student can take O/L by school maximum 2 times"
Call modform.ClearTextBoxes(Me)
Exit Sub
End If

If (dcAesthetic.Text = dcAddtional1.Text Or dcAesthetic.Text = dcAddtional2.Text Or dcAddtional1.Text = dcAddtional2.Text) Then
MsgBox "Check the Subjects"
Exit Sub
End If



RecOlRes.AddNew
RecOlRes!stuid = Trim(dcSID.Text)
RecOlRes!IndexNo = Text2.Text
RecOlRes!OlYear = Text3.Text

RecOlRes!Maths = Trim(Text4.Text)
RecOlRes!Science = Trim(Text5.Text)
RecOlRes!Social = Trim(Text6.Text)
RecOlRes!English = Trim(Text7.Text)

RecOlRes!Language = Trim(dcLanguage.Text) + " " + Trim(Text8.Text)
RecOlRes!Religion = Trim(dcReligion.Text) + " " + Trim(Text9.Text)
RecOlRes!Aesthetic = Trim(dcAesthetic.Text) + " " + Trim(Text10.Text)
RecOlRes!Technical = Trim(dcTechnical.Text) + " " + Trim(Text11.Text)
If (Text12.Text = "") Then
RecOlRes!Additional1 = ""
Else
RecOlRes!Additional1 = Trim(dcAddtional1.Text) + " " + Trim(Text12.Text)
End If
If (Trim(Text13.Text) = "") Then
RecOlRes!Additional2 = ""
Else
RecOlRes!Additional2 = Trim(dcAddtional2.Text) + " " + Trim(Text13.Text)
End If


RecOlRes.UpdateBatch
RecOlRes.Requery



End Sub

Private Sub dcSID_Change()
Dim sid As String
sid = Trim(dcSID.Text)
    If sid <> "" Then
    RecOlStu.MoveFirst
    RecOlStu.Find "StuID = '" & sid & "'"
      If RecOlStu.EOF Then
            Text1.Text = "Invalid Student ID"
            Command2.Enabled = False
        Else
            Text1.Text = RecOlStu!StudentName
            Command2.Enabled = True
            Set RecOlRes = openDB.OpenRecord("select * from OLRESULT where StuId='" & sid & "'")
       End If
     Else
     Text1.Text = ""
        
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "O/L Result")
    
    Set RecOlStu = openDB.OpenRecord("select A.StuID,M.StudentName from ACTIVESTUDENT A,MAINSTUDENTS M WHERE A.StuID=M.StuID and A.Curr_Class LIKE '11%'")
    'Set RecOlRes = openDB.OpenRecord("select * from OLRESULT")

    Set RecAddtionl = openDB.OpenRecord("select * from OLSUBJECT where Category='Additional Subjects'")
    Set RecAest = openDB.OpenRecord("select * from OLSUBJECT where Category='Aesthetic Subjects'")
    Set RecLang = openDB.OpenRecord("select * from OLSUBJECT where Category='Core Subjects' and SubjectNames LIKE '%Language%'")
    Set RecReligions = openDB.OpenRecord("select * from OLSUBJECT where Category='Religions'")
    Set RecTech = openDB.OpenRecord("select * from OLSUBJECT where Category IN ('Commerce Stream','Home Economics Stream','Technical Stream','Technical Subjects / Agriculture Stream')")


    RecOlStu.MoveFirst
    dcSID.ListField = "StuID"
    Set dcSID.RowSource = RecOlStu
    
    RecLang.MoveFirst
    dcLanguage.ListField = "SubjectNames"
    Set dcLanguage.RowSource = RecLang
    
    RecReligions.MoveFirst
    dcReligion.ListField = "SubjectNames"
    Set dcReligion.RowSource = RecReligions
    
    RecAddtionl.MoveFirst
    dcAddtional1.ListField = "SubjectNames"
    Set dcAddtional1.RowSource = RecAddtionl
    
    RecAddtionl.MoveFirst
    dcAddtional2.ListField = "SubjectNames"
    Set dcAddtional2.RowSource = RecAddtionl
    
    RecAest.MoveFirst
    dcAesthetic.ListField = "SubjectNames"
    Set dcAesthetic.RowSource = RecAest
    
    RecTech.MoveFirst
    dcTechnical.ListField = "SubjectNames"
    Set dcTechnical.RowSource = RecTech
    
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecOlStu.Close
RecOlRes.Close
RecAddtionl.Close
RecAest.Close
RecLang.Close
RecReligions.Close
RecTech.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Text4.SetFocus
Else
    msg = MsgBox("Year should be numeric value", vbExclamation)
    KeyAscii = 0
End If

End Sub
