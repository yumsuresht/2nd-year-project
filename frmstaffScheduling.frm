VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmstaffScheduling 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   Icon            =   "frmstaffScheduling.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6705
   Begin VB.CommandButton Command2 
      Height          =   315
      Left            =   3480
      Picture         =   "frmstaffScheduling.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Search"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   6000
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgschedule 
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6165
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPeriod 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin MSDataListLib.DataCombo dcSubject 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      IntegralHeight  =   0   'False
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   2295
   End
   Begin MSDataListLib.DataCombo dcStaffID 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcClass 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Frame Frame1 
      Caption         =   "Class Teacher"
      Height          =   1095
      Left            =   4200
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
      Begin VB.CheckBox Check1 
         Caption         =   "Class Teacher"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label5 
      Caption         =   "No_Of_Period"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Subject"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Class"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Staff ID"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmstaffScheduling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStaff As ADODB.Recordset
Dim RecClass As ADODB.Recordset
Dim RecSubject As ADODB.Recordset
Dim RecStaffSchedule As ADODB.Recordset


Private Sub cmdclear_Click()
Call modform.ClearTextBoxes(Me)
dcClass.Text = ""
dcStaffID.Text = ""
dcSubject.Text = ""

End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
RecStaffSchedule.Delete
RecStaffSchedule.UpdateBatch
RecStaffSchedule.Requery

End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
Dim classte As String
If (Check1.Value = 1) Then
classte = "Yes"
Else
classte = "No"
End If

RecStaffSchedule!Curr_Year = Year(Date)
RecStaffSchedule!ClassName = Trim(dcClass)
RecStaffSchedule!staffid = Trim(dcStaffID)
RecStaffSchedule!SubjectNames = Trim(dcSubject)
RecStaffSchedule!Period = Val(txtPeriod.Text)
RecStaffSchedule!ClassTeacher = Trim(classte)
RecStaffSchedule.UpdateBatch
RecStaffSchedule.Requery

If (Err.Number <> 0) Then
MsgBox "Already add this record"
Exit Sub
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmStaffSearch.Show
modform.formname = "StaffSch"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim classte As String
If (Trim(dcClass.Text) = "" Or Trim(dcStaffID.Text) = "" Or Trim(dcSubject.Text) = "" Or Trim(txtPeriod.Text) = "") Then
MsgBox "Blank fields"
Exit Sub
End If


If (Check1.Value = 1) Then
classte = "Yes"
Else
classte = "No"
End If

RecStaffSchedule.AddNew
RecStaffSchedule!Curr_Year = Year(Date)
RecStaffSchedule!ClassName = Trim(dcClass.Text)
RecStaffSchedule!staffid = Trim(dcStaffID.Text)
RecStaffSchedule!SubjectNames = Trim(dcSubject.Text)
RecStaffSchedule!Period = Val(txtPeriod.Text)
RecStaffSchedule!ClassTeacher = Trim(classte)
RecStaffSchedule.UpdateBatch
RecStaffSchedule.Requery
Check1.Value = 0

If (Err.Number <> 0) Then
MsgBox "Already add this record"
Exit Sub
End If

End Sub

Private Sub Command4_Click()
MsgBox Val(Left((dcClass.Text), 2))
End Sub

Private Sub dcClass_Change()
On Error Resume Next
Dim grade As Integer
grade = Val(Left((dcClass.Text), 2))
If (grade = 12 Or grade = 13) Then
    Set RecSubject = openDB.OpenRecord("select distinct(SubjectNames) from ALSUBJECT")
Else
    Set RecSubject = openDB.OpenRecord("select distinct(SubjectNames) from SUBJECT")
End If

    dcSubject.ListField = "SubjectNames"
    Set dcSubject.RowSource = RecSubject
    dcSubject.Text = UCase("")
End Sub

Private Sub dcStaffID_Change()
Dim staffid As String
staffid = Trim(dcStaffID.Text)
    If staffid <> "" Then
    RecStaff.MoveFirst
    RecStaff.Find "StaffID = '" & staffid & "'"
      If RecStaff.EOF Then
            txtname.Text = "Invalid StaffID"
            Command3.Enabled = False
        Else
            txtname.Text = RecStaff!FullName
            Command3.Enabled = True
        End If
        
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "Teacher Scheduling")
    Set RecStaff = openDB.OpenRecord("select StaffID,FullName from STAFF")
    Set RecClass = openDB.OpenRecord("select * from CLASS")
    Set RecSubject = openDB.OpenRecord("select distinct(SubjectNames) from SUBJECT union select distinct(SubjectNames) from ALSUBJECT")
    Set RecStaffSchedule = openDB.OpenRecord("select * from TEACHERSCHEDULE")
    
    dcClass.ListField = "ClassName"
    Set dcClass.RowSource = RecClass
    dcClass.Text = ""

    dcSubject.ListField = UCase("SubjectNames")
    Set dcSubject.RowSource = RecSubject
    dcSubject.Text = UCase("")

    
    dcStaffID.ListField = "StaffID"
    Set dcStaffID.RowSource = RecStaff
    dcStaffID.Text = ""


    RecStaffSchedule.MoveFirst
    Set dgschedule.DataSource = RecStaffSchedule
    Command3.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecStaff.Close
RecClass.Close
RecSubject.Close
RecStaffSchedule.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub
