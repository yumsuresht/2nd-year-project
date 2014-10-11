VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmnewstaff 
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   Icon            =   "frmnewstaff.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   10320
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   9615
      Left            =   12600
      TabIndex        =   50
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         Height          =   375
         Left            =   360
         TabIndex        =   58
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   360
         TabIndex        =   57
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Search"
         Height          =   375
         Left            =   360
         TabIndex        =   56
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   55
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   54
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New"
         Height          =   375
         Left            =   360
         TabIndex        =   53
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<<"
         Height          =   375
         Left            =   360
         TabIndex        =   52
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   ">>"
         Height          =   375
         Left            =   1320
         TabIndex        =   51
         Top             =   2400
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      Begin MSMask.MaskEdBox Text10 
         Height          =   345
         Left            =   8040
         TabIndex        =   49
         Top             =   5160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   11
         Mask            =   "###-#######"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text9 
         Height          =   345
         Left            =   8040
         TabIndex        =   48
         Top             =   4560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   11
         Mask            =   "###-#######"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame4 
         Caption         =   "Working"
         Height          =   855
         Left            =   6240
         TabIndex        =   45
         Top             =   5640
         Width           =   4215
         Begin VB.OptionButton Option6 
            Caption         =   "Temporary Staff"
            Height          =   375
            Left            =   2280
            TabIndex        =   47
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Permanent Staff"
            Height          =   375
            Left            =   720
            TabIndex        =   46
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   8040
         MaxLength       =   3
         TabIndex        =   44
         Text            =   "20"
         Top             =   6720
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dcpost 
         Height          =   315
         Left            =   1560
         TabIndex        =   42
         Top             =   2280
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sex"
         Height          =   735
         Left            =   240
         TabIndex        =   20
         Top             =   3480
         Width           =   3855
         Begin VB.OptionButton Option1 
            Caption         =   "Male"
            Height          =   255
            Left            =   720
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Female"
            Height          =   255
            Left            =   2400
            TabIndex        =   21
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Civil Status"
         Height          =   735
         Left            =   240
         TabIndex        =   17
         Top             =   4560
         Width           =   3855
         Begin VB.OptionButton Option3 
            Caption         =   "Single"
            Height          =   255
            Left            =   720
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Married"
            Height          =   255
            Left            =   2400
            TabIndex        =   18
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   2880
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmnewstaff.frx":030A
         Left            =   8040
         List            =   "frmnewstaff.frx":031D
         TabIndex        =   12
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#########a"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmnewstaff.frx":0366
         Left            =   8040
         List            =   "frmnewstaff.frx":0376
         TabIndex        =   9
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   8040
         TabIndex        =   8
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   8040
         TabIndex        =   7
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   7680
         Width           =   2295
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   6360
         Width           =   2295
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   8280
         Width           =   2295
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   8880
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   6960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/M/yyyy"
         Format          =   45350915
         UpDown          =   -1  'True
         CurrentDate     =   38287
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8040
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/M/yyyy"
         Format          =   45350915
         UpDown          =   -1  'True
         CurrentDate     =   38287
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
         Left            =   10440
         TabIndex        =   60
         Top             =   9240
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "(Years)"
         Height          =   255
         Left            =   3480
         TabIndex        =   59
         Top             =   9000
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "* Working Hours"
         Height          =   255
         Left            =   6240
         TabIndex        =   43
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "StaffID"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "* Name with Initial"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "* RegNo"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "PostHeld"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "* Grade"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Nationality"
         Height          =   255
         Left            =   6240
         TabIndex        =   36
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Date of Birth"
         Height          =   255
         Left            =   6240
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "NIC No"
         Height          =   255
         Left            =   6240
         TabIndex        =   34
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Religion"
         Height          =   255
         Left            =   6240
         TabIndex        =   33
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Street"
         Height          =   255
         Left            =   6240
         TabIndex        =   32
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "City"
         Height          =   255
         Left            =   6240
         TabIndex        =   31
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Private Tel Number"
         Height          =   255
         Left            =   6240
         TabIndex        =   30
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Permanant Tel Number"
         Height          =   255
         Left            =   6240
         TabIndex        =   29
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "W_O_P Number"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "FileNo"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   7680
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Qualification"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Date of Appointment"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "* Salary"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   8280
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Service"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   8880
         Width           =   1215
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
         Left            =   10300
         TabIndex        =   61
         Top             =   9200
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmnewstaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Dim Rec2 As ADODB.Recordset
Dim Rec3, Rec4 As ADODB.Recordset
Dim Reclogin As ADODB.Recordset
Dim RecMedLeaves As ADODB.Recordset
Dim RecCasLeaves As ADODB.Recordset


Dim sex, civ, work As String


Private Sub Command1_Click()
    Call modform.ClearTextBoxes(Me)
Command2.Caption = "New"

End Sub

Private Sub Command12_Click()
frmStaffSearch1.Show
Command2.Caption = "New"

End Sub

Private Sub Command14_Click()
Unload Me

End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim na As String
If (Command2.Caption = "New") Then
Call modform.ClearTextBoxes(Me)
Command2.Caption = "Update"
Text1.Text = rec1!STAFFIDS + 1
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False

Else
 Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True

If (Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text14.Text = "") Then
MsgBox "Please fill the mandatory fields"
Exit Sub
End If

If (Val(Text4.Text) > 30) Then
MsgBox "Working hours between (0-30)"
Exit Sub
End If
Command2.Caption = "New"
   




Call modform.uppercase(Me)

If Option1.Value = True Then
sex = UCase(Option1.Caption)
ElseIf Option2.Value = True Then
sex = UCase(Option2.Caption)
End If

If Option3.Value = True Then
civ = UCase(Option3.Caption)
ElseIf Option4.Value = True Then
civ = UCase(Option4.Caption)
End If


If Option5.Value = True Then
work = UCase("Permanent")
ElseIf Option6.Value = True Then
work = UCase("Temporary")
End If

If Option1.Value = True Then
na = "MR"
ElseIf Option2.Value = True Then
    If Option3.Value = True Then
    na = "MISS"
    ElseIf Option4.Value = True Then
    na = "MRS"
    End If
End If




 rec.MoveFirst
        rec.AddNew
        rec!staffid = rec1!STAFFIDS + 1
        rec!FullName = na + " " + Text2.Text
        rec!RegNo = Text3.Text
        rec!PostHeld = dcpost.Text
        rec!grade = Text5.Text
        rec!Nationality = Combo1.Text
        rec!D_Of_B = DTPicker1.Value
        rec!NIC = Text6.Text
        rec!sex = sex
        rec!Religion = Combo2.Text
        rec!CivilStatus = civ
        rec!Street = Text7.Text
        rec!City = Text8.Text
        rec!Pri_Tp = Text9.Text
        rec!Per_Tp = Text10.Text
        rec!WOP = Text11.Text
        rec!FileNo = Text12.Text
        rec!Qualification = Text13.Text
        rec!D_Of_Appionment = DTPicker2.Value
        rec!Salary = Text14.Text
        rec!Service = Text15.Text
        rec!Category = work
        rec!Work_Hours = Text4.Text
        rec.UpdateBatch
        rec.Requery
        Rec2.Requery
        
        Reclogin.AddNew
        Reclogin!staffid = rec1!STAFFIDS + 1
        Reclogin!Passwords = rec1!STAFFIDS + 1
        Reclogin.UpdateBatch
        Reclogin.Requery
        
        If Option5.Value = True Then
        RecMedLeaves.AddNew
        RecMedLeaves!staffid = rec1!STAFFIDS + 1
        RecMedLeaves!CurrentYear = 21
        RecMedLeaves.UpdateBatch
        RecMedLeaves.Requery
        
        RecCasLeaves.AddNew
        RecCasLeaves!staffid = rec1!STAFFIDS + 1
        RecCasLeaves!TotalDays = 20
        RecCasLeaves.UpdateBatch
        RecCasLeaves.Requery
        End If
        
        rec1!STAFFIDS = rec1!STAFFIDS + 1
        rec1.UpdateBatch
        
        
        
        
        If (Err.Number = 0) Then
        MsgBox "Sucessfully Added"
        End If
        
    main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
    Call modform.ClearTextBoxes(Me)
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim stid As String

If MsgBox("Are you sure do you want to delete the current record?", vbYesNo, "Delete") = vbYes Then

Rec3.AddNew
Rec3!LeStId = rec!staffid
Rec3!FullName = rec!FullName
Rec3!RegNo = rec!RegNo
Rec3!PostHeld = rec!PostHeld
Rec3!grade = rec!grade
Rec3!Nationality = rec!Nationality
Rec3!D_Of_B = rec!D_Of_B
Rec3!NIC = rec!NIC
Rec3!sex = rec!sex
Rec3!Religion = rec!Religion
Rec3!CivilStatus = rec!CivilStatus
Rec3!Street = rec!Street
Rec3!City = rec!City
Rec3!Pri_Tp = rec!Pri_Tp
Rec3!Per_Tp = rec!Per_Tp
Rec3!WOP = rec!WOP
Rec3!FileNo = rec!FileNo
Rec3!Qualification = rec!Qualification
Rec3!Service = rec!Service
Rec3!Category = rec!Category
Rec3!D_Of_Appionment = rec!D_Of_Appionment
Rec3.UpdateBatch
Rec3.Requery

stid = Trim(rec!staffid)
      
    If stid <> "" Then
        Rec4.MoveFirst
                Rec4.Find "SCID = '" & stid & "'"
                If Rec4.EOF Then
                                  
                Else
                    Rec4.Delete
                    Rec4.UpdateBatch
                    Rec4.Requery
                End If
    End If


rec.Delete
rec.UpdateBatch
rec.Requery
Rec2.Requery
display
End If


End Sub

Private Sub Command4_Click()
On Error Resume Next
rec.MovePrevious
    If rec.BOF Then rec.MoveFirst
    display
Command3.Enabled = True
Command5.Enabled = True
Command2.Caption = "New"

End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim aaa
Text1.Locked = True
If (Text1.Text = "") Then
MsgBox "Blank fields"
Exit Sub
End If

Call modform.uppercase(Me)

If Option1.Value = True Then
sex = UCase(Option1.Caption)
ElseIf Option2.Value = True Then
sex = UCase(Option2.Caption)
End If

If Option3.Value = True Then
civ = UCase(Option3.Caption)
ElseIf Option4.Value = True Then
civ = UCase(Option4.Caption)
End If

If Option5.Value = True Then
work = UCase("Permanent")
ElseIf Option6.Value = True Then
work = UCase("Temporary")
End If
        If (MsgBox("Do you really want to edit this current record?", vbYesNo, "Delete") = vbYes) Then
        rec!FullName = Text2.Text
        rec!RegNo = Text3.Text
        rec!PostHeld = dcpost.Text
        rec!grade = Text5.Text
        rec!Nationality = Combo1.Text
        rec!D_Of_B = DTPicker1.Value
        rec!NIC = Text6.Text
        rec!sex = sex
        rec!Religion = Combo2.Text
        rec!CivilStatus = civ
        rec!Street = Text7.Text
        rec!City = Text8.Text
        rec!Pri_Tp = Text9.Text
        rec!Per_Tp = Text10.Text
        rec!WOP = Text11.Text
        rec!FileNo = Text12.Text
        rec!Qualification = Text13.Text
        rec!D_Of_Appionment = DTPicker2.Value
        rec!Salary = Text14.Text
        rec!Service = Text15.Text
        rec!Category = work
        rec!Work_Hours = Text4.Text
        
        rec.UpdateBatch
        rec.Requery
        Rec2.Requery
         MsgBox "Record changed"
        End If
        
        
        
    main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
    Call modform.ClearTextBoxes(Me)
    Command2.Caption = "New"

End Sub

Private Sub Command6_Click()
On Error Resume Next
rec.MoveNext
    If rec.EOF Then rec.MoveLast
    display
    
Command3.Enabled = True
Command5.Enabled = True
Command2.Caption = "New"

End Sub

Private Sub Command7_Click()
End Sub

Private Sub Command8_Click()
Dim s As String
s = modform.Staff1
rec.MoveFirst
rec.Find "StaffID = '" & s & "'"
    If rec.EOF Then
     MsgBox "Cannot find"
     Exit Sub
    Else
      display
    End If

End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "New Staff")
    Set rec = openDB.OpenRecord("SELECT * FROM STAFF ")
    Set rec1 = openDB.OpenRecord("SELECT * FROM IDS")
    Set Rec2 = openDB.OpenRecord("SELECT DISTINCT(PostHeld) FROM STAFF")
    Set Rec3 = openDB.OpenRecord("SELECT * FROM LEAVINGSTAFF")
    Set Rec4 = openDB.OpenRecord("select * from LIBRARYMEMBER")
    Set Reclogin = openDB.OpenRecord("select * from LOGIN")
    Set RecMedLeaves = openDB.OpenRecord("select * from MEDICALLEAVES")
    Set RecCasLeaves = openDB.OpenRecord("select * from CASUALLEAVES")

    dcpost.ListField = "PostHeld"
    Set dcpost.RowSource = Rec2
    
   If (Val(rec1!STAFFIDS) = 0) Then
   MsgBox "You must initilized the staffid"
   
    frminitial.Show
    frminitial.Text3.SetFocus
    frminitial.Text3.BackColor = &H80000018

    Unload Me
   Exit Sub
   End If
    Call Command4_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
rec.Close
rec1.Close
Rec2.Close
Rec3.Close
Rec4.Close
Reclogin.Close
RecMedLeaves.Close
RecCasLeaves.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Public Sub display()
        Text1.Text = rec!staffid
        Text2.Text = rec!FullName
        Text3.Text = rec!RegNo
        dcpost.Text = rec!PostHeld
        Text5.Text = rec!grade
        Combo1.Text = rec!Nationality
        DTPicker1.Value = rec!D_Of_B
        Text6.Text = rec!NIC
        
        If (rec!sex = "MALE") Then Option1.Value = True
        If (rec!sex = "FEMALE") Then Option2.Value = True
        
        If (Trim(rec!CivilStatus) = "SINGLE") Then Option3.Value = True
        If (Trim(rec!CivilStatus) = "MARRIED") Then Option4.Value = True
        
        If (Trim(rec!Category) = "TEMPORARY") Then Option6.Value = True
        If (Trim(rec!Category) = "PERMANENT") Then Option5.Value = True

    
        Combo2.Text = rec!Religion
        Text7.Text = rec!Street
        Text8.Text = rec!City
        Text9.Text = rec!Pri_Tp
        Text10.Text = rec!Per_Tp
        Text11.Text = rec!WOP
        Text12.Text = rec!FileNo
        Text13.Text = rec!Qualification
        DTPicker2.Value = rec!D_Of_Appionment
        Text14.Text = rec!Salary
        Text15.Text = rec!Service
        Text4.Text = rec!Work_Hours
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Text15.SetFocus
Else
    msg = MsgBox("Salary should be numeric value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Combo1.SetFocus
Else
    msg = MsgBox("Year should be numeric value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Command2.SetFocus
Else
    msg = MsgBox("Working hours should be numeric value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Or KeyAscii = 88 Or KeyAscii = 86 Or KeyAscii = 118 Or KeyAscii = 120 Then
ElseIf KeyAscii = 13 Then
    Combo2.SetFocus
    If (Len(Text6.Text) <> 10) Then
    MsgBox "Check your NIC"
    Text6.SetFocus
    Exit Sub
    ElseIf (Right(Text6.Text, 1) <> "V" And Right(Text6.Text, 1) <> "v" And Right(Text6.Text, 1) <> "X" And Right(Text6.Text, 1) <> "x") Then
    MsgBox "Check your NIC"
    Text6.SetFocus
    Exit Sub
    End If
    
Else
    msg = MsgBox("Invalid NIC No", vbExclamation)
    KeyAscii = 0
End If
End Sub

Public Sub filllists()
Dim s As String
s = modform.Staff1
rec.MoveFirst
rec.Find "StaffID = '" & s & "'"
    If rec.EOF Then
     MsgBox "Cannot find"
     Exit Sub
    Else
      display
    End If
End Sub

