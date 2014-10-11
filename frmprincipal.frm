VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmprincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   Icon            =   "frmprincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6285
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Attendance"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3975
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7011
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
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Class"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No of Students"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Leave Update"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdclassupdate 
      Caption         =   "Class Update"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Class and Students"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStudent As ADODB.Recordset
Dim RecTermAvg As ADODB.Recordset
Dim RecYearAvg As ADODB.Recordset
Dim RecMedical As ADODB.Recordset
Dim RecStudent1 As ADODB.Recordset



Private Sub Cancel_Click()
Unload Me
End Sub


Private Sub cmdclassupdate_Click()
On Error Resume Next
Dim grade As Integer
Dim class As String
Dim YearAvg As Double
Dim Years As Integer
Dim ans
ans = MsgBox("Are you sure to update class", vbOKCancel)

If (ans = vbOK) Then

RecTermAvg.MoveFirst
RecYearAvg.MoveFirst
While Not RecTermAvg.EOF
    RecYearAvg.MoveFirst
    RecYearAvg.Find "StuID = '" & RecTermAvg!stuid & "'"
    If ((RecTermAvg!grade <> "11 R1") And (RecTermAvg!grade <> "11 R2") And (RecTermAvg!grade <> "Leave") And (RecTermAvg!grade <> "13 R1") And (RecTermAvg!grade <> "13 R2")) Then
        YearAvg = (RecTermAvg!Term1 + RecTermAvg!Term2 + RecTermAvg!Term3) / 3
        Years = Trim(Left(RecTermAvg!grade, 2))
        RecYearAvg.Fields(Years - 5) = YearAvg
    End If
    RecTermAvg.MoveNext
Wend
RecYearAvg.UpdateBatch
RecYearAvg.Requery



RecStudent.MoveFirst
While Not RecStudent.EOF
If (RecStudent!Curr_Class <> "Leave") Then

    grade = Trim(Left(RecStudent!Curr_Class, 2))
    class = Trim(Right(RecStudent!Curr_Class, 2))
    
    
    If (grade < 11) Then
        RecStudent!Curr_Class = ((Val(grade) + 1) & " " & class)
    ElseIf (grade = 11) Then
        If (class = "R1") Then
            RecStudent!Curr_Class = Trim("11 R2")
        ElseIf (class = "R2") Then
            RecStudent!Curr_Class = Trim("Leave")
        Else
            RecStudent!Curr_Class = Trim("11 R1")
        End If
    ElseIf (11 < grade) And (grade < 13) Then
         RecStudent!Curr_Class = ((Val(grade) + 1) & " " & class)
    ElseIf (grade = 13) Then
        If (class = "R1") Then
            RecStudent!Curr_Class = Trim("13 R2")
        ElseIf (class = "R2") Then
            RecStudent!Curr_Class = Trim("Leave")
        Else
            RecStudent!Curr_Class = Trim("13 R1")
        End If
    End If
End If
     RecStudent.MoveNext
    
      
    
    
Wend
RecStudent.UpdateBatch
RecStudent.Requery


RecStudent.MoveFirst 'Class
RecTermAvg.MoveFirst 'average

While Not RecStudent.EOF
    RecTermAvg.MoveFirst
    RecTermAvg.Find "StuID = '" & RecStudent!stuid & "'"
    RecTermAvg!grade = RecStudent!Curr_Class
    RecStudent.MoveNext
Wend
RecTermAvg.UpdateBatch
RecTermAvg.Requery


RecTermAvg.MoveFirst
While Not RecTermAvg.EOF
    RecTermAvg!Term1 = 0
    RecTermAvg!Term2 = 0
    RecTermAvg!Term3 = 0
    RecTermAvg.MoveNext
Wend
RecTermAvg.UpdateBatch
RecTermAvg.Requery

Call filllist

MsgBox "All student class are updated sucessful"
End If

End Sub



Private Sub Command1_Click()
On Error Resume Next
Dim ans1
ans1 = MsgBox("Are you sure to update records", vbOKCancel)
If (ans1 = vbOK) Then
RecMedical.MoveFirst
While Not RecMedical.EOF
    
    RecMedical!TwoYearsBefore = RecMedical!OneYearBefore
    RecMedical!OneYearBefore = RecMedical!CurrentYear
    RecMedical!CurrentYear = 21
    RecMedical.MoveNext
Wend
RecMedical.UpdateBatch
RecMedical.Requery
MsgBox "updated sucessful"
End If

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()
frmDateSelect.Show

End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "Return Book")

Set RecStudent = openDB.OpenRecord("select * from ACTIVESTUDENT")
Set RecTermAvg = openDB.OpenRecord("SELECT StuID,Grade,Term1,Term2,Term3 FROM TERMAVG ")
Set RecYearAvg = openDB.OpenRecord("SELECT * FROM YEARAVERAGE ")
Set RecMedical = openDB.OpenRecord("SELECT * FROM MEDICALLEAVES ")
Set RecStudent = openDB.OpenRecord("select Curr_Class,count(*) from ACTIVESTUDENT group by(Curr_Class)")
Call filllist

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecStudent.Close
RecTermAvg.Close
RecMedical.Close
RecStudent1.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Public Sub filllist()
On Error Resume Next
RecStudent.MoveFirst
If Not (RecStudent.EOF And RecStudent.BOF) Then
    RecStudent.MoveFirst
        lv1.ListItems.clear
        While Not RecStudent.EOF
            Set List = lv1.ListItems.add
            List.Text = RecStudent.Fields(0)
            List.SubItems(1) = RecStudent.Fields(1)
            RecStudent.MoveNext
        Wend
        Else
        lv1.ListItems.clear
    End If
End Sub
