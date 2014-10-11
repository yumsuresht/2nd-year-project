VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMarks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   9435
   Begin MSFlexGridLib.MSFlexGrid ms1 
      Height          =   6735
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11880
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      Begin MSDataListLib.DataCombo dcclass 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Class"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   7920
      Width           =   1215
   End
End
Attribute VB_Name = "frmMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStu As ADODB.Recordset
Dim RecClass As ADODB.Recordset
Dim RecGradeSub As ADODB.Recordset






Private Sub Command1_Click()
Unload Me
End Sub

Private Sub dcclass_Change()
On Error Resume Next
Dim class As String
Dim co As Integer
class = Trim(Left(dcclass.Text, 2))
    ms1.Cols = 2
ms1.Rows = 1
Set RecGradeSub = openDB.OpenRecord("select TotalSubjects, Subject1,Subject2,Subject3,Subject4,Subject5,Subject6,Subject7,Subject8,Subject9 , Subject10, Subject11, Subject12 from GRADEANDSUBJECTS where Grade='" & class & "'")
Set RecStu = openDB.OpenRecord("select A.StuID,M.StudentName,M.FatherName,A.Curr_Class from ACTIVESTUDENT A,MAINSTUDENTS M where A.StuID=M.StuID and A.Curr_Class='" & dcclass.Text & "'")
co = RecStu.RecordCount
ms1.Rows = 100
    ms1.Cols = RecGradeSub!TotalSubjects + 2
    ms1.ColWidth(0) = 1000
    ms1.ColWidth(1) = 2000
    
    
    ms1.ColWidth(2) = 700
    ms1.ColWidth(3) = 700
    ms1.ColWidth(4) = 700
    ms1.ColWidth(5) = 700
    ms1.ColWidth(6) = 700
    ms1.ColWidth(7) = 700
    ms1.ColWidth(8) = 700
    ms1.ColWidth(9) = 700
    ms1.ColWidth(10) = 700
    ms1.ColWidth(11) = 700
    ms1.ColWidth(12) = 700
    ms1.ColWidth(13) = 700
    

    
    ms1.TextMatrix(0, 0) = "StdID"
    ms1.TextMatrix(0, 1) = "Student Name"
    ms1.TextMatrix(0, 2) = RecGradeSub!Subject1
    ms1.TextMatrix(0, 3) = RecGradeSub!Subject2
    ms1.TextMatrix(0, 4) = RecGradeSub!Subject3
    ms1.TextMatrix(0, 5) = RecGradeSub!Subject4
    ms1.TextMatrix(0, 6) = RecGradeSub!Subject5
    ms1.TextMatrix(0, 7) = RecGradeSub!Subject6
    ms1.TextMatrix(0, 8) = RecGradeSub!Subject7
    ms1.TextMatrix(0, 9) = RecGradeSub!Subject8
    ms1.TextMatrix(0, 10) = RecGradeSub!Subject9
    ms1.TextMatrix(0, 11) = RecGradeSub!Subject10
    ms1.TextMatrix(0, 12) = RecGradeSub!Subject11
    ms1.TextMatrix(0, 13) = RecGradeSub!Subject12
    
  
        


Dim i As Integer
i = 1
If Not (RecStu.EOF And RecStu.BOF) Then
    RecStu.MoveFirst
        
        While Not RecStu.EOF
            ms1.TextMatrix(i, 0) = RecStu!stuid
            ms1.TextMatrix(i, 1) = RecStu!StudentName
            i = i + 1
            RecStu.MoveNext
        Wend
        
End If











End Sub

Private Sub Form_Load()
On Error Resume Next
Dim a, a1 As String
    Call modform.FormSize(Me, "Student Marks")
    Set RecClass = openDB.OpenRecord("select * from Class order by ClassName")




dcclass.ListField = "ClassName"
Set dcclass.RowSource = RecClass
dcclass.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecGradeSub.Close
RecClass.Close
RecStu.Close

End Sub
