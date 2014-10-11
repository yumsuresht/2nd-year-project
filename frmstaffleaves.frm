VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmstaffleaves 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   Icon            =   "frmstaffleaves.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6720
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   24
         Top             =   3480
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   3480
         Picture         =   "frmstaffleaves.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Search"
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4200
         TabIndex        =   19
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/M/yyyy"
         Format          =   177405955
         CurrentDate     =   38276
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/M/yyyy"
         Format          =   177405955
         CurrentDate     =   38276
      End
      Begin VB.TextBox txtAviCas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4440
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   1560
         TabIndex        =   11
         Top             =   1080
         Width           =   3855
         Begin VB.OptionButton Option3 
            Caption         =   "Duty"
            Height          =   255
            Left            =   2880
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Medical"
            Height          =   255
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Casual"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.TextBox txtDes 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   1560
         TabIndex        =   7
         Top             =   2520
         Width           =   4695
      End
      Begin VB.TextBox txtAviMed 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   3855
      End
      Begin MSDataListLib.DataCombo dcStaffID 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label7 
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
         Left            =   5040
         TabIndex        =   25
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "No of Leave Days"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000A&
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Available Medical Leave"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Availabe Casual Leave"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Description"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "StaffID"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "frmstaffleaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStaff As ADODB.Recordset
Dim RecMed As ADODB.Recordset
Dim RecCas As ADODB.Recordset
Dim RecLea As ADODB.Recordset

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim leavetypes As String
Dim message As String
If (txtname.Text = "") Then
MsgBox "Invalid Staff ID"
Exit Sub
End If

If (Option1.Value = True) Then
leavetypes = "Casual"
RecCas!UseDays = Val(RecCas!UseDays) + Val(Text1.Text)
RecCas.UpdateBatch
RecCas.Requery
End If




If (Option2.Value = True) Then
leavetypes = "Medical"
    If (Val(RecMed!CurrentYear) > 0) Then
        RecMed!CurrentYear = Val(RecMed!CurrentYear) - Val(Text1.Text)
    ElseIf (Val(RecMed!CurrentYear) = 0 And Val(RecMed!OneYearBefore) > 0) Then
        RecMed!OneYearBefore = Val(RecMed!OneYearBefore) - Val(Text1.Text)
    ElseIf (Val(RecMed!CurrentYear) = 0 And Val(RecMed!OneYearBefore) = 0 And Val(RecMed!TwoYearsBefore) > 0) Then
        RecMed!TwoYearsBefore = Val(RecMed!TwoYearsBefore) - Val(Text1.Text)
    ElseIf (Val(RecMed!CurrentYear) = 0 And Val(RecMed!OneYearBefore) = 0 And Val(RecMed!TwoYearsBefore) = 0) Then
        message = MsgBox("Your Medical leaves are finished. Other leaves are calculate as no pay leave. Click ok to proceed", vbOKCancel)
        If (message = vbOK) Then
            RecMed!Nopay = Val(RecMed!Nopay) + Val(Text1.Text)
        Else
            Exit Sub
        End If
    End If

End If
RecMed.UpdateBatch
RecMed.Requery



RecLea.AddNew
RecLea!staffid = dcStaffID.Text
RecLea!Dateto = DTPicker1.Value
RecLea!DateFrom = DTPicker2.Value
RecLea!LeaveType = leavetypes
RecLea!LeaveDetails = txtDes.Text
RecLea.UpdateBatch
RecLea.Requery


Call Form_Load
Call modform.ClearTextBoxes(Me)
End Sub

Private Sub Command2_Click()

frmStaffSearch.Show
modform.formname = "StaffLeave"

End Sub

Private Sub dcStaffID_Change()
On Error Resume Next
Dim stafid As String
stafid = dcStaffID.Text
    If stafid <> "" Then
        RecStaff.MoveFirst
        RecMed.MoveFirst
        RecCas.MoveFirst
        Label5.Caption = "(Current Year)"
        RecStaff.Find "StaffID = '" & stafid & "'"
        RecMed.Find "StaffID = '" & stafid & "'"
        RecCas.Find "StaffID = '" & stafid & "'"
        
        If RecStaff.EOF Then
            txtname.Text = ""
        Else
            txtname.Text = RecStaff!FullName
        End If
        
        If RecMed.EOF Then
            txtAviMed.Text = ""
        Else
            If (Val(RecMed!CurrentYear) > 0) Then
                txtAviMed.Text = RecMed!CurrentYear
                Label5.Caption = "(Current Year)"
            ElseIf (Val(RecMed!CurrentYear) = 0 And Val(RecMed!OneYearBefore) > 0) Then
                txtAviMed.Text = RecMed!OneYearBefore
                Label5.Caption = "(Last Year)"
            ElseIf (Val(RecMed!CurrentYear) = 0 And Val(RecMed!OneYearBefore) = 0 And Val(RecMed!TwoYearsBefore) > 0) Then
                txtAviMed.Text = RecMed!TwoYearsBefore
                Label5.Caption = "(One Year Before)"
            ElseIf (Val(RecMed!CurrentYear) = 0 And Val(RecMed!OneYearBefore) = 0 And Val(RecMed!TwoYearsBefore) = 0) Then
                txtAviMed.Text = RecMed!Nopay
                Label5.Caption = "(Number of No Pay Leaves)"
            End If
        End If
        
        If RecCas.EOF Then
            txtAviCas.Text = ""
        Else
            txtAviCas.Text = Val(Val(RecCas!TotalDays) - Val(RecCas!UseDays))
        End If
        
    End If
End Sub

Private Sub DTPicker1_Change()
DTPicker2.Value = DTPicker1.Value + 1
End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "Staff Leaves")
Set RecStaff = openDB.OpenRecord("select * from STAFF ")
Set RecMed = openDB.OpenRecord("select * from MEDICALLEAVES ")
Set RecCas = openDB.OpenRecord("select * from CASUALLEAVES ")
Set RecLea = openDB.OpenRecord("select * from STAFFLEAVES ")

DTPicker1.Value = Date
DTPicker2.Value = Date + 1

Label7.Caption = Format(Date, "dd/mm/yyyy")

dcStaffID.ListField = "StaffID"
Set dcStaffID.RowSource = RecStaff
dcStaffID.Text = ""




End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecStaff.Close
RecMed.Close
RecCas.Close
RecLea.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Timer1_Timer()
 Text1.Text = IIf(DateDiff("d", CDate(DTPicker1.Value), CDate(DTPicker2.Value)) > 0, DateDiff("d", CDate(DTPicker1.Value), CDate(DTPicker2.Value)), 0)
End Sub
