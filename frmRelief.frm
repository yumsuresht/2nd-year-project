VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelief 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmRelief.frx":0000
   LinkTopic       =   "frmRelief"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvAbs 
      Height          =   5175
      Left            =   6720
      TabIndex        =   22
      Top             =   720
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   9128
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Staff ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Staff Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Class"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Teach Subject"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   8400
      Top             =   6360
   End
   Begin VB.Frame Frame1 
      Caption         =   "SUBJECT TIME INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3375
      Left            =   10440
      TabIndex        =   2
      Top             =   6000
      Width           =   4575
      Begin VB.Label Label5 
         Caption         =   "01:20:00 PM  TO  02:00:00 PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   18
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "12:40:00 PM  TO  01:20:00 PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   1680
         TabIndex        =   17
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "12:00:00 PM  TO  12:40:00 PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   1680
         TabIndex        =   16
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "11:20:00 AM  TO  12:00:00 PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   15
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "10:10:00 AM  TO  10:50:00 AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   14
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "09:30:00 AM  TO  10:10:00 AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   13
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "08:50:00 AM  TO  09:30:00 AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   12
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "08:00:00 AM  TO  08:50:00 AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "PERIOD 7"
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
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "PERIOD 8"
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
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "PERIOD 6"
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
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "PERIOD 5"
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
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "PERIOD 4"
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
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "PERIOD 3"
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
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "PERIOD 2"
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "PERIOD 1"
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
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4200
      Top             =   9480
   End
   Begin MSComctlLib.ListView lvabsstaff 
      Height          =   8655
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   15266
      View            =   3
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
         Text            =   "Staff ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full Name"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Label Label9 
      Caption         =   "Total absent Staffs today"
      Height          =   375
      Left            =   6840
      TabIndex        =   24
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   9120
      TabIndex        =   23
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000011&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6720
      TabIndex        =   21
      Top             =   6840
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000011&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Top             =   6480
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000011&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   6120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Not Available Staff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmRelief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecRelife As ADODB.Recordset
Dim RecTimetable As ADODB.Recordset



Private Sub Command2_Click()
Text1.Text = displayPeriod
End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "Relief Work")
Call relife
Label6.Caption = "PERIOD " & displayPeriod
Label2.Caption = Time
Label7.Caption = Date
End Sub
Public Sub filllist(rec As ADODB.Recordset)
If Not (rec.EOF And rec.BOF) Then
    rec.MoveFirst
        lvabsstaff.ListItems.clear
        While Not rec.EOF
            Set List = lvabsstaff.ListItems.add
            List.Text = rec.Fields(0)
            List.SubItems(1) = rec.Fields(1)
            rec.MoveNext
        Wend
        Else
        lvabsstaff.ListItems.clear
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecTimetable.Close
RecRelife.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Timer1_Timer()
Call relife


End Sub

Public Sub filllist1(rec1 As ADODB.Recordset)
If Not (rec1.EOF And rec1.BOF) Then
    rec1.MoveFirst
        lvAbs.ListItems.clear
        While Not rec1.EOF
            Set List = lvAbs.ListItems.add
            List.Text = rec1.Fields(0)
            List.SubItems(1) = rec1.Fields(1)
            List.SubItems(2) = rec1.Fields(2)
            List.SubItems(3) = rec1.Fields(3)
            rec1.MoveNext
        Wend
        Else
        lvAbs.ListItems.clear
    End If
End Sub


Public Sub relife()
On Error Resume Next
Dim sql As String
Dim sql1 As String
sql = "select StaffID,FullName from STAFF where PostHeld IN ('TEACHER','SECTIONAL HEAD','DEPUTY PRINCIPAL') and StaffID NOT IN(" & " select distinct(staffID) from STAFFATTENDANCE where day(attDate)=day(GETDATE()) and  month(attDate)=month(GETDATE()) and  year(attDate)=year(GETDATE()) and GoingTime IS NULL and  NOT EXISTS (select StaffID from SHORTLEAVES where day(dates)=day(GETDATE()) and month(dates)=month(GETDATE()) and year(dates)=year(GETDATE()) and ComingTime IS NULL))"
Set RecRelife = openDB.OpenRecord(sql)

'Set dgsub.DataSource = RecTimetable
sql1 = "select S.StaffID,S.FullName,T.ClassName,T.SubjectNames from TIMETABLE T , STAFF S where S.STAFFID=T.STAFFID and  Days='" & Format(Now, "dddd") & "' and Period='" & displayPeriod & "' and T.STAFFID  not IN(select distinct(staffID)From StaffAttendance Where Day(attDate) = Day(GETDATE()) and  month(attDate)=month(GETDATE()) and  year(attDate)=year(GETDATE()) and GoingTime IS NULL and NOT EXISTS (select StaffID from SHORTLEAVES Where Day(dates) = Day(GETDATE()) and month(dates)=month(GETDATE()) and year(dates)=year(GETDATE()) and ComingTime IS NULL))"
Set RecTimetable = openDB.OpenRecord(sql1)
'Set dgsub.DataSource = RecTimetable


Call filllist(RecRelife)
Call filllist1(RecTimetable)
Label8.Caption = RecRelife.RecordCount

End Sub

Public Function displayPeriod() As Integer
Dim s As Integer
Dim a As Date
a = Format(Time, "HH:MM:SS")
If (a > "08:00:00 AM" And a <= "08:50:00 AM") Then
s = 1
ElseIf (a > "08:50:00 AM" And a <= "09:30:00 AM") Then
s = 2
ElseIf (a > "09:30:31 AM" And a <= "10:10:31 AM") Then
s = 3
ElseIf (a > "10:10:00 AM" And a <= "10:50:00 AM") Then
s = 4
ElseIf (a > "10:50:00 AM" And a <= "11:19:00 AM") Then
s = 11
ElseIf (a > "11:20:00 AM" And a <= "12:00:00 PM") Then
s = 5
ElseIf (a > "12:00:00 PM" And a <= "12:40:00 PM") Then
s = 6
ElseIf (a > "12:40:00 PM" And a <= "13:20:00 PM") Then
s = 7
ElseIf (a > "13:20:00 PM" And a <= "14:00:00 PM") Then
s = 8
End If
displayPeriod = s
End Function

Private Sub Timer2_Timer()
If (displayPeriod = 11) Then
Label6.Caption = "Interval Time"
Else
Label6.Caption = "PERIOD " & displayPeriod
End If
Label2.Caption = Time
Label7.Caption = Date
End Sub
