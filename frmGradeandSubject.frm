VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGradeandSubject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11625
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   8760
      TabIndex        =   54
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10200
      TabIndex        =   51
      Top             =   6120
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvclasssub 
      Height          =   5655
      Left            =   6240
      TabIndex        =   3
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   25
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Grade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Period/Week-Subj 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Subject 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Period/Week-Subj 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Subject 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Period/Week-Subj 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Subject 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Period/Week-Subj 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Subject 5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Period/Week-Subj 5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Subject 6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Period/Week-Subj 6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Subject 7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Period/Week-Subj 7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Subject 8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Period/Week-Subj 8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Subject 9"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Period/Week-Subj 9"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Subject 10"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Period/Week-Subj 10"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Subject 11"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Period/Week-Subj 11"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Subject 12"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Period/Week-Subj 12"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Subjects"
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6015
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   49
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   47
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   45
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   43
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   41
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   39
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   37
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   35
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   33
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   31
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   315
         Left            =   1200
         TabIndex        =   18
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo5 
         Height          =   315
         Left            =   1200
         TabIndex        =   20
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo6 
         Height          =   315
         Left            =   1200
         TabIndex        =   21
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo7 
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo8 
         Height          =   315
         Left            =   1200
         TabIndex        =   23
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo9 
         Height          =   315
         Left            =   1200
         TabIndex        =   24
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo10 
         Height          =   315
         Left            =   1200
         TabIndex        =   25
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo11 
         Height          =   315
         Left            =   1200
         TabIndex        =   26
         Top             =   3600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo12 
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   3960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo13 
         Height          =   315
         Left            =   1200
         TabIndex        =   28
         Top             =   4320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   315
         Left            =   1200
         TabIndex        =   52
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label25 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   50
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label24 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   48
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label23 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   46
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   44
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   42
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   40
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   38
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   36
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label17 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "No of Period/Week"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grade"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   56
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox grades 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Total No.of Subjects"
         Height          =   255
         Left            =   3360
         TabIndex        =   55
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Grade"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmGradeandSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecGrade As ADODB.Recordset
Dim RecSubject As ADODB.Recordset
Dim RecGradeSub As ADODB.Recordset



Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next

RecGradeSub.AddNew
RecGradeSub!grade = grades.Text
RecGradeSub!TotalSubjects = Text13.Text

RecGradeSub!Subject1 = DataCombo2.Text
RecGradeSub!No_of_Per_Sub1 = Text1.Text

RecGradeSub!Subject2 = DataCombo3.Text
RecGradeSub!No_of_Per_Sub2 = Text2.Text

RecGradeSub!Subject3 = DataCombo4.Text
RecGradeSub!No_of_Per_Sub3 = Text3.Text

RecGradeSub!Subject4 = DataCombo5.Text
RecGradeSub!No_of_Per_Sub4 = Text4.Text

RecGradeSub!Subject5 = DataCombo6.Text
RecGradeSub!No_of_Per_Sub5 = Text5.Text

RecGradeSub!Subject6 = DataCombo7.Text
RecGradeSub!No_of_Per_Sub6 = Text6.Text

RecGradeSub!Subject7 = DataCombo8.Text
RecGradeSub!No_of_Per_Sub7 = Text7.Text

RecGradeSub!Subject8 = DataCombo9.Text
RecGradeSub!No_of_Per_Sub8 = Text8.Text

RecGradeSub!Subject9 = DataCombo10.Text
RecGradeSub!No_of_Per_Sub9 = Text9.Text

RecGradeSub!Subject10 = DataCombo11.Text
RecGradeSub!No_of_Per_Sub10 = Text10.Text

RecGradeSub!Subject11 = DataCombo12.Text
RecGradeSub!No_of_Per_Sub11 = Text11.Text

RecGradeSub!Subject12 = DataCombo13.Text
RecGradeSub!No_of_Per_Sub12 = Text12.Text


RecGradeSub.UpdateBatch
RecGradeSub.Requery

Call filllist(RecGradeSub)


End Sub

Private Sub Form_Load()
On Error Resume Next
Dim st1 As Integer
Dim st2 As Integer

    Call modform.FormSize(Me, "GRADE AND SUBJECTS")
    Set RecGrade = openDB.OpenRecord("select * from school")
    Set RecGradeSub = openDB.OpenRecord("select * from GRADEANDSUBJECTS")

    Call filllist(RecGradeSub)
    

    st1 = RecGrade!StartGrade
    st2 = RecGrade!EndGrade
    grades.clear
    
    For i = st1 To st2
        grades.AddItem i
    Next i
    

       
        


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecGrade.Close
RecSubject.Close
RecGradeSub.Close
End Sub



Private Sub grades_Click()
On Error Resume Next
Dim grade As Integer
grade = Val(grades.Text)
If (grade = 12 Or grade = 13) Then
    Set RecSubject = openDB.OpenRecord("select distinct(SubjectNames) from ALSUBJECT")
Else
    Set RecSubject = openDB.OpenRecord("select distinct(SubjectNames) from SUBJECT")
End If


    DataCombo2.ListField = "SubjectNames"
    Set DataCombo2.RowSource = RecSubject
    DataCombo2.Text = UCase("")
    
    DataCombo3.ListField = "SubjectNames"
    Set DataCombo3.RowSource = RecSubject
    DataCombo3.Text = UCase("")
    
    DataCombo4.ListField = "SubjectNames"
    Set DataCombo4.RowSource = RecSubject
    DataCombo4.Text = UCase("")
    
    DataCombo5.ListField = "SubjectNames"
    Set DataCombo5.RowSource = RecSubject
    DataCombo5.Text = UCase("")
    
    DataCombo6.ListField = "SubjectNames"
    Set DataCombo6.RowSource = RecSubject
    DataCombo6.Text = UCase("")
    
    DataCombo7.ListField = "SubjectNames"
    Set DataCombo7.RowSource = RecSubject
    DataCombo7.Text = UCase("")
    
    DataCombo8.ListField = "SubjectNames"
    Set DataCombo8.RowSource = RecSubject
    DataCombo8.Text = UCase("")
    
    DataCombo9.ListField = "SubjectNames"
    Set DataCombo9.RowSource = RecSubject
    DataCombo9.Text = UCase("")
    
    DataCombo10.ListField = "SubjectNames"
    Set DataCombo10.RowSource = RecSubject
    DataCombo10.Text = UCase("")
    
    DataCombo11.ListField = "SubjectNames"
    Set DataCombo11.RowSource = RecSubject
    DataCombo11.Text = UCase("")
    
    DataCombo12.ListField = "SubjectNames"
    Set DataCombo12.RowSource = RecSubject
    DataCombo12.Text = UCase("")
    
    DataCombo13.ListField = "SubjectNames"
    Set DataCombo13.RowSource = RecSubject
    DataCombo13.Text = UCase("")
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
'
    Dim strValid As String
    '
    strValid = "0123456789+-."
    '
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Public Sub filllist(Rec5 As ADODB.Recordset)

On Error Resume Next
    
    
If Not (Rec5.EOF And Rec5.BOF) Then
    Rec5.MoveFirst
        lvclasssub.ListItems.clear
        While Not Rec5.EOF
            Set List = lvclasssub.ListItems.add
            List.Text = Rec5.Fields(0)
            List.SubItems(1) = Rec5.Fields(2)
            List.SubItems(2) = Rec5.Fields(3)
            List.SubItems(3) = Rec5.Fields(4)
            List.SubItems(4) = Rec5.Fields(5)
            List.SubItems(5) = Rec5.Fields(6)
            List.SubItems(6) = Rec5.Fields(7)
            List.SubItems(7) = Rec5.Fields(8)
            List.SubItems(8) = Rec5.Fields(9)
            List.SubItems(9) = Rec5.Fields(10)
            List.SubItems(10) = Rec5.Fields(11)
            List.SubItems(11) = Rec5.Fields(12)
            List.SubItems(12) = Rec5.Fields(13)
            List.SubItems(13) = Rec5.Fields(14)
            List.SubItems(14) = Rec5.Fields(15)
            List.SubItems(15) = Rec5.Fields(16)
            List.SubItems(16) = Rec5.Fields(17)
            List.SubItems(17) = Rec5.Fields(18)
            List.SubItems(18) = Rec5.Fields(19)
            List.SubItems(19) = Rec5.Fields(20)
            List.SubItems(20) = Rec5.Fields(21)
            List.SubItems(21) = Rec5.Fields(22)
            List.SubItems(22) = Rec5.Fields(23)
            List.SubItems(23) = Rec5.Fields(24)
            List.SubItems(24) = Rec5.Fields(24)
            Rec5.MoveNext
        Wend
        Else
        lvclasssub.ListItems.clear
End If

End Sub

