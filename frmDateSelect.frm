VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDateSelect 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin MSComCtl2.DTPicker date1 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/M/yyyy"
         Format          =   75563011
         CurrentDate     =   38336
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/M/yyyy"
         Format          =   75563011
         CurrentDate     =   38336
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmDateSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecAttendance As ADODB.Recordset
Dim d1, d2, s1 As String
Dim rpl


Private Sub Command1_Click()
On Error Resume Next
If (date1.Value > date2.Value) Then
MsgBox "Invalid Dates, Please check"
Exit Sub
End If
s1 = "The daily attendance details between " & Format(date1.Value, "dd-mm-yyyy") & " to " & Format(date1.Value, "dd-mm-yyyy") & " will be delete. Are you sure?"
rpl = MsgBox(s1, vbOKCancel)
If (rpl = vbOK) Then
While RecAttendance.RecordCount <> 0
    
    RecAttendance.Delete
    RecAttendance.UpdateBatch
    RecAttendance.Requery
Wend
Unload Me

Else
Unload Me
End If



End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub date1_Change()
date2.Value = date1.Value + 1
d1 = Format(date1.Value, "yyyy-mm-dd")
d2 = Format(date2.Value, "yyyy-mm-dd")
Set RecAttendance = openDB.OpenRecord("select * from STAFFATTENDANCE where AttDate between '" & d1 & "'and '" & d2 & "'")
Label3.Caption = "Total Records  :" & RecAttendance.RecordCount

End Sub

Private Sub date2_Change()
d1 = Format(date1.Value, "yyyy-mm-dd")
d2 = Format(date2.Value, "yyyy-mm-dd")
Set RecAttendance = openDB.OpenRecord("select * from STAFFATTENDANCE where AttDate between '" & d1 & "'and '" & d2 & "'")
Label3.Caption = "Total Records  :" & RecAttendance.RecordCount

End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "Date Selecter")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecAttendance.Close
End Sub
