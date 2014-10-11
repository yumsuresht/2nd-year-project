VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7110
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   15
      Left            =   1560
      TabIndex        =   47
      Top             =   3960
      Width           =   5295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   4200
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   120
         Top             =   3960
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3135
         Left            =   2880
         TabIndex        =   46
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5530
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmWizard.frx":030A
      End
      Begin VB.Frame Frame4 
         Caption         =   "Study Days"
         Height          =   3255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   2415
         Begin VB.CheckBox Check7 
            Caption         =   "Sunday"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Saturday"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Friday"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Thursday"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Wednessday"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Tuesday"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Monday"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Finish"
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "< Back"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Descriptions"
         Height          =   255
         Left            =   2880
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2400
         TabIndex        =   49
         Top             =   2640
         Width           =   855
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "frmWizard.frx":038C
         Left            =   3840
         List            =   "frmWizard.frx":03B7
         TabIndex        =   36
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmWizard.frx":03E6
         Left            =   2400
         List            =   "frmWizard.frx":0411
         TabIndex        =   34
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Height          =   315
         Left            =   2400
         TabIndex        =   32
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmWizard.frx":0440
         Left            =   3840
         List            =   "frmWizard.frx":044A
         TabIndex        =   29
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   2400
         TabIndex        =   28
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   2400
         TabIndex        =   27
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmWizard.frx":0456
         Left            =   3840
         List            =   "frmWizard.frx":0460
         TabIndex        =   25
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmWizard.frx":046C
         Left            =   3840
         List            =   "frmWizard.frx":0476
         TabIndex        =   24
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   2400
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   2400
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "< Back"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Next >"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "[Minutes]"
         Height          =   255
         Left            =   3360
         TabIndex        =   50
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Number of Class Rooms"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "To"
         Height          =   255
         Left            =   3360
         TabIndex        =   35
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Grade"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Class / Day"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Period Hour"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Interval Time"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "School End Time"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "School Start Time"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         Top             =   3240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/M/yyyy"
         Format          =   19726339
         CurrentDate     =   38322
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmWizard.frx":0482
         Left            =   1560
         List            =   "frmWizard.frx":048F
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   1560
         TabIndex        =   13
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next >"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "School Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Type of School"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "WebSite"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Name of School"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   4800
      Left            =   0
      Picture         =   "frmWizard.frx":04BC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1440
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecSchool As ADODB.Recordset
Private Sub Command1_Click()
Frame2.Visible = True
Frame1.Visible = False
End Sub

Private Sub Command2_Click()
Frame2.Visible = True
Frame3.Visible = False
End Sub

Private Sub Command3_Click()
On Error Resume Next
If (Text1.Text = "" Or Text2.Text = "") Then
MsgBox "Check your school informations"
Exit Sub
End If
If (RecSchool.RecordCount = 0) Then
Call AddDatas
Else
Call UpdateDatas
End If

Unload Me
End Sub

Private Sub Command4_Click()
Frame3.Visible = True
Frame2.Visible = False
End Sub

Private Sub Command5_Click()
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call modform.FormSize(Me, "Information Wizard")
Set RecSchool = openDB.OpenRecord("select * from SCHOOL")
main.Enabled = False
Call ShowDatas
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecSchool.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Command4.SetFocus
Else
    KeyAscii = 0
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change to uppercase

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Text9.SetFocus
Else
    KeyAscii = 0
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

Private Sub Timer1_Timer()
If (Frame1.Visible = True) Then
frmWizard.Caption = "Wizard 1 of 3"
ElseIf (Frame2.Visible = True) Then
frmWizard.Caption = "Wizard 2 of 3"
ElseIf (Frame3.Visible = True) Then
frmWizard.Caption = "Wizard 3 of 3"
End If
End Sub

Public Function AddDatas()
On Error Resume Next
RecSchool.AddNew
RecSchool!SCHOOLNAME = Text1.Text
RecSchool!ADDRESS1 = Text2.Text
RecSchool!ADDRESS2 = Text3.Text
RecSchool!WEB = Text4.Text
RecSchool!Type = Combo1.Text
RecSchool!STARTDATE = DTPicker1.Value

RecSchool!STARTTIME = Text5.Text + " " + Combo2.Text
RecSchool!ENDTIME = Text6.Text + " " + Combo3.Text
RecSchool!Interval = Text7.Text + " " + Combo4.Text
RecSchool!PERIODHOUR = Text8.Text
RecSchool!NUMPERIOD = Text9.Text
RecSchool!StartGrade = Combo5.Text
RecSchool!EndGrade = Combo6.Text
RecSchool!NOCLASSROOM = Text10.Text

RecSchool!MONDAY = Check1.Value
RecSchool!TUESDAY = Check2.Value
RecSchool!WEDNESSDAY = Check3.Value
RecSchool!THURSDAY = Check4.Value
RecSchool!FRIDAY = Check5.Value
RecSchool!SATURDAY = Check6.Value
RecSchool!SUNDAY = Check7.Value
RecSchool!DESCRIPTIONS = RichTextBox1.Text

RecSchool.UpdateBatch
RecSchool.Requery

End Function

Public Function UpdateDatas()
On Error Resume Next
Dim ans
ans = MsgBox("Are you sure to save this records?", vbOKCancel)
If (ans = vbOK) Then

RecSchool!SCHOOLNAME = Text1.Text
RecSchool!ADDRESS1 = Text2.Text
RecSchool!ADDRESS2 = Text3.Text
RecSchool!WEB = Text4.Text
RecSchool!Type = Combo1.Text
RecSchool!STARTDATE = DTPicker1.Value

RecSchool!STARTTIME = Text5.Text + " " + Combo2.Text
RecSchool!ENDTIME = Text6.Text + " " + Combo3.Text
RecSchool!Interval = Text7.Text + " " + Combo4.Text
RecSchool!PERIODHOUR = Text8.Text
RecSchool!NUMPERIOD = Text9.Text
RecSchool!StartGrade = Combo5.Text
RecSchool!EndGrade = Combo6.Text
RecSchool!NOCLASSROOM = Text10.Text


RecSchool!MONDAY = Check1.Value
RecSchool!TUESDAY = Check2.Value
RecSchool!WEDNESSDAY = Check3.Value
RecSchool!THURSDAY = Check4.Value
RecSchool!FRIDAY = Check5.Value
RecSchool!SATURDAY = Check6.Value
RecSchool!SUNDAY = Check7.Value
RecSchool!DESCRIPTIONS = RichTextBox1.Text

RecSchool.UpdateBatch
RecSchool.Requery
End If

End Function

Public Function ShowDatas()
On Error Resume Next
Text1.Text = RecSchool!SCHOOLNAME
Text2.Text = RecSchool!ADDRESS1
Text3.Text = RecSchool!ADDRESS2
Text4.Text = RecSchool!WEB
Combo1.Text = RecSchool!Type
DTPicker1.Value = RecSchool!STARTDATE

Text5.Text = Left(RecSchool!STARTTIME, Len(RecSchool!STARTTIME) - 3)
Combo2.Text = Right(RecSchool!STARTTIME, 2)
Text6.Text = Left(RecSchool!ENDTIME, Len(RecSchool!ENDTIME) - 3)
Combo3.Text = Right(RecSchool!ENDTIME, 2)
Text7.Text = Left(RecSchool!Interval, Len(RecSchool!Interval) - 3)
Combo4.Text = Right(RecSchool!Interval, 2)


Text8.Text = RecSchool!PERIODHOUR
Text9.Text = RecSchool!NUMPERIOD
Combo5.Text = RecSchool!StartGrade
Combo6.Text = RecSchool!EndGrade
Text10.Text = RecSchool!NOCLASSROOM


Check1.Value = ConBool(RecSchool!MONDAY)
Check2.Value = ConBool(RecSchool!TUESDAY)
Check3.Value = ConBool(RecSchool!WEDNESSDAY)
Check4.Value = ConBool(RecSchool!THURSDAY)
Check5.Value = ConBool(RecSchool!FRIDAY)
Check6.Value = ConBool(RecSchool!SATURDAY)
Check7.Value = ConBool(RecSchool!SUNDAY)


RichTextBox1.Text = RecSchool!DESCRIPTIONS
End Function

Public Function ConBool(ans As Boolean) As Integer
If (ans = True) Then
ConBool = 1
Else
ConBool = 0
End If
End Function
