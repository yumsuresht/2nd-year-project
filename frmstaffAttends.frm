VERSION 5.00
Begin VB.Form frmstaffAttends 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   Icon            =   "frmstaffAttends.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7680
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   4440
      TabIndex        =   11
      Top             =   720
      Width           =   3135
      Begin VB.Label Label8 
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "DATE : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "TIME :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5400
      Top             =   1440
   End
   Begin VB.Frame Frame2 
      Caption         =   "Leave"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "Duty"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Short"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "IN"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "OUT"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Short leave in this month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4215
      Begin VB.CommandButton Command3 
         Caption         =   "OUT"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "IN"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label10 
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
      Left            =   4440
      TabIndex        =   19
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Staffs can't take short leave  between 8.00 a.m to 9.00 a.m and 1.00 p.m to 2.00 p.m"
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
      Height          =   615
      Left            =   4440
      TabIndex        =   18
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label5 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmstaffAttends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStfAtt As ADODB.Recordset
Dim RecStfAtt1 As ADODB.Recordset
Dim RecStaff As ADODB.Recordset
Dim ResShortLe As ADODB.Recordset
Dim ResShortLe1 As ADODB.Recordset
Dim Reccoshort As ADODB.Recordset
Dim leave As String
Dim userid As String


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
RecStfAtt.AddNew
RecStfAtt!staffid = userid
RecStfAtt!attDate = Date
RecStfAtt!ComingTime = Time
RecStfAtt.UpdateBatch
RecStfAtt.Requery
Command2.Enabled = False
Form_Load
End Sub

Private Sub Command3_Click()
On Error Resume Next
RecStfAtt1!GoingTime = Time
RecStfAtt1.UpdateBatch
RecStfAtt1.Requery
Command3.Enabled = False
Form_Load

End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim opt As String
Dim reply
If (Option1.Value = True) Then
opt = "Short"
ElseIf (Option2.Value = True) Then
opt = "Duty"
End If

If (Val(Label2.Caption) > 1) And (Option2.Value = False) Then
    reply = MsgBox("You can take two short leave with in a month. Addtional short leaves are calculate as half day leave", vbOKCancel)
    If (reply = vbOK) Then
        ResShortLe.AddNew
        ResShortLe!staffid = userid
        ResShortLe!Dates = Date
        ResShortLe!GoingTime = Now
        ResShortLe!LeaveType = opt
        ResShortLe.UpdateBatch
        ResShortLe.Requery
        Command4.Enabled = False
        Form_Load
    Else
        Command3.Enabled = True
        Exit Sub
    End If
Else
    ResShortLe.AddNew
    ResShortLe!staffid = userid
    ResShortLe!Dates = Date
    ResShortLe!GoingTime = Now
    ResShortLe!LeaveType = opt
    ResShortLe.UpdateBatch
    ResShortLe.Requery
    Command4.Enabled = False
    Form_Load
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
ResShortLe!ComingTime = Now
ResShortLe.UpdateBatch
ResShortLe.Requery
Command5.Enabled = False
Form_Load

End Sub

Private Sub Form_Load()
On Error Resume Next
userid = modform.userid
Dim currhour As Integer
currhour = Hour(Time)
Call modform.FormSize(Me, "Staff Attendance")
Set RecStfAtt = openDB.OpenRecord("select * from STAFFATTENDANCE where AttDate='" & Format(Date, "yyyy-mm-dd") & "' and StaffID='" & userid & "'")
Set RecStfAtt1 = openDB.OpenRecord("select * from STAFFATTENDANCE where goingtime is null and AttDate='" & Format(Date, "yyyy-mm-dd") & "' and StaffID='" & userid & "'")
Set RecStaff = openDB.OpenRecord("select * from STAFF where StaffID='" & userid & "'")
Set ResShortLe = openDB.OpenRecord("select * from SHORTLEAVES where Dates='" & Format(Date, "yyyy-mm-dd") & "' and StaffID='" & userid & "'")
Set ResShortLe1 = openDB.OpenRecord("select * from SHORTLEAVES where ComingTime is null and Dates='" & Format(Date, "yyyy-mm-dd") & "' and StaffID='" & userid & "'")

If (RecStfAtt.RecordCount = 0) And (currhour < 14) Then
Command2.Enabled = True
Else
Command2.Enabled = False
End If

If (RecStfAtt1.RecordCount = 1) Then
Command3.Enabled = True
Else
Command3.Enabled = False
End If

If (ResShortLe.RecordCount = 0) Then
Command4.Enabled = True
Else
Command4.Enabled = False
End If

If (ResShortLe1.RecordCount = 1) Then
Command5.Enabled = True
Else
Command5.Enabled = False
End If

If (Command2.Enabled = True) Then
Command4.Enabled = False
End If

If (Command3.Enabled = False) Then
Command4.Enabled = False
End If

If (Command4.Enabled = False) Then
Command3.Enabled = False
End If

If (Command5.Enabled = False) Then
Command3.Enabled = True
End If

If (RecStfAtt1.RecordCount = 1) Then
Command3.Enabled = True
Else
Command3.Enabled = False
End If

If (Command5.Enabled = True) Then
Command3.Enabled = False
End If

If Not ((currhour >= 9) And (currhour < 13)) Then
Command4.Enabled = False
Command5.Enabled = False
End If

Set Reccoshort = openDB.OpenRecord("select * from SHORTLEAVES where month(dates)=month(GETDATE()) and LeaveType ='Short' and StaffID='" & userid & "'")
Label2.Caption = Reccoshort.RecordCount

Label6.Caption = RecStaff!FullName
Label8.Caption = Format(Date, "dd-mmmm yyyy")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecStfAtt.Close
RecStfAtt1.Close
RecStaff.Close
ResShortLe.Close
ResShortLe1.Close
Reccoshort.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Timer1_Timer()

Label4.Caption = Time
End Sub
