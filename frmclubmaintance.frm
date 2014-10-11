VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmclubmaintance 
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   10560
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Caption         =   "Maintain"
      Height          =   855
      Left            =   10440
      TabIndex        =   38
      Top             =   240
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3240
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "New Member"
         Height          =   375
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add"
         Height          =   375
         Left            =   1800
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Height          =   8295
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   15015
      Begin MSComctlLib.ListView lvmain 
         Height          =   3975
         Left            =   7560
         TabIndex        =   35
         Top             =   4200
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Year"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Supervisior ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Supervisior Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "President StudentID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "President Student Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Vice President StudentID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Vice President Student Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5400
         TabIndex        =   34
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd - MMM - yyyy"
         Format          =   46006275
         UpDown          =   -1  'True
         CurrentDate     =   38245
      End
      Begin MSComctlLib.ListView lvclubmem 
         Height          =   3975
         Left            =   120
         TabIndex        =   32
         Top             =   4200
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "StuID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Student Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Join Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Frame Frame7 
         Caption         =   "New Member"
         Height          =   3615
         Left            =   9480
         TabIndex        =   21
         Top             =   240
         Width           =   5415
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   375
            Left            =   1320
            TabIndex        =   25
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   375
            Left            =   1320
            TabIndex        =   24
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Height          =   975
            Left            =   1320
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   1800
            Width           =   3855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Add"
            Height          =   375
            Left            =   4080
            TabIndex        =   22
            Top             =   2880
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dcclubmem 
            Height          =   315
            Left            =   1320
            TabIndex        =   26
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16776960
            Text            =   ""
         End
         Begin VB.Label Label8 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Student ID"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Join Date"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Description"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "(mm/dd/yyyy)"
            Height          =   255
            Left            =   3480
            TabIndex        =   27
            Top             =   1440
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Descriptions"
         Height          =   3015
         Left            =   4800
         TabIndex        =   19
         Top             =   840
         Width           =   4575
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   2655
            Left            =   120
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   4300
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Supervisor"
         Height          =   1215
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   4575
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1080
            TabIndex        =   15
            Top             =   720
            Width           =   3375
         End
         Begin MSDataListLib.DataCombo dcstaff 
            Height          =   315
            Left            =   1080
            TabIndex        =   16
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label2 
            Caption         =   "Staff ID"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Vice President"
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   4575
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1080
            TabIndex        =   10
            Top             =   840
            Width           =   3375
         End
         Begin MSDataListLib.DataCombo dcstu2 
            Height          =   315
            Left            =   1080
            TabIndex        =   11
            Top             =   360
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label6 
            Caption         =   "Student ID"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "President"
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   4575
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1080
            TabIndex        =   5
            Top             =   720
            Width           =   3375
         End
         Begin MSDataListLib.DataCombo dcstu1 
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label4 
            Caption         =   "Student ID"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Label Label13 
         Caption         =   "Annual Details"
         Height          =   255
         Left            =   10440
         TabIndex        =   37
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Member 
         Caption         =   "Member Details"
         Height          =   255
         Left            =   2880
         TabIndex        =   36
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Year 
         Caption         =   "Date"
         Height          =   255
         Left            =   4920
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CLUBS/UNIONS"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin MSDataListLib.DataCombo dcclubname 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         Caption         =   "Name"
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
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmclubmaintance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecStaff As ADODB.Recordset
Public RecStu As ADODB.Recordset
Public Recclub As ADODB.Recordset
Public Recclubmain As ADODB.Recordset
Public Recclubmain1 As ADODB.Recordset
Public Recclubmem As ADODB.Recordset
Public Recclubmem1 As ADODB.Recordset



Private Sub Command1_Click()
Unload Me
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub DataCombo4_Click(Area As Integer)

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
On Error Resume Next
If (Trim(dcclubname.Text) = "" Or Trim(dcclubmem.Text) = "") Then
    MsgBox "Blankfields"
    Exit Sub
End If

If (MsgBox("Are you sure?", vbOKCancel) = vbOK) Then
    Recclubmem.AddNew
    Recclubmem!CName = Trim(dcclubname.Text)
    Recclubmem!stuid = Trim(dcclubmem.Text)
    Recclubmem!JoinDate = Text6.Text
    Recclubmem!DESCRIPTIONS = Text7.Text
    Recclubmem.UpdateBatch
    Recclubmem.Requery
    If (Err.Number <> 0) Then
        MsgBox "This student already registered"
        Exit Sub
    End If
Else
Exit Sub
End If
dcclubname_Change
Command5.Caption = "New Member"
        
End Sub

Private Sub Command5_Click()
If (Command5.Caption = "New Member") Then
Command5.Caption = "Hide"
Frame7.Visible = True
Text6.Text = Date
ElseIf (Command5.Caption = "Hide") Then
Command5.Caption = "New Member"
Frame7.Visible = False
End If

End Sub

Private Sub Command6_Click()
On Error Resume Next
If (Trim(Text2.Text) = "" Or Trim(Text3.Text) = "" Or Trim(Text4.Text) = "" Or Trim(dcclubname.Text) = "") Then
MsgBox "Blankfields"
Exit Sub
ElseIf (Trim(dcstu1.Text) = Trim(dcstu2.Text)) Then
MsgBox "Duplicate StudentID"
Exit Sub
Else

        Recclubmain.AddNew
        Recclubmain!CName = Trim(dcclubname.Text)
        Recclubmain!Years = DTPicker1.Value
        Recclubmain!Super_StaffID = Trim(dcStaff.Text)
        Recclubmain!Super_StafName = Trim(Text2.Text)
        Recclubmain!Pres_StuID = Trim(dcstu1.Text)
        Recclubmain!Pres_StuName = Trim(Text3.Text)
        Recclubmain!Sec_StuID = Trim(dcstu2.Text)
        Recclubmain!Sec_StuName = Trim(Text4.Text)
        Recclubmain!DESCRIPTIONS = Trim(Text1.Text)
        
        Recclubmain.UpdateBatch
        Recclubmain.Requery
        
    If (Err.Number <> 0) Then
        MsgBox "This student already registered"
        Exit Sub
    End If
End If

dcclubname_Change
End Sub

Private Sub dcclubmem_Change()
On Error Resume Next
Dim stuid As String
   
    stuid = Trim(dcclubmem.Text)
    
           If stuid <> "" Then
                RecStu.MoveFirst
                RecStu.Find "StuID = '" & stuid & "'"
                If RecStu.EOF Then
                    Text5.Text = ""
                Else
                    Text5.Text = RecStu!FatherName + " " + RecStu!StudentName
                    
               End If
            End If
End Sub

Private Sub dcclubname_Change()
On Error Resume Next
Dim clubname As String
clubname = Trim(dcclubname.Text)
    If clubname <> "" Then
        Recclub.MoveFirst
        Recclub.Find "CName = '" & clubname & "'"
        
        If Recclub.EOF Then
            Frame6.Visible = False
            Command5.Enabled = False
            Command6.Enabled = False
        Else
            Frame6.Visible = True
            Frame7.Visible = False
            Command5.Enabled = True
            Command6.Enabled = True
            Recclubmain1.Close
            Set Recclubmem1 = openDB.OpenRecord("select * from MAINSTUDENTS M,ACTIVESTUDENT A,CLUBMEMBER c WHERE M.StuID=A.StuID AND A.StuID=C.StuID and C.CName='" + Trim(dcclubname.Text) + "'")
            Set Recclubmain1 = openDB.OpenRecord("SELECT * FROM CLUBMAINTAINCE where CName='" + Trim(dcclubname.Text) + "'")
            
            dcstu1.Text = ""
            dcstu2.Text = ""
            dcStaff.Text = ""
            dcclubmem.Text = ""
            Call modform.ClearTextBoxes(Me)
            dcstu1.Refresh
            dcstu2.Refresh
                   
                   
            dcstu1.ListField = "StuID"
            Set dcstu1.RowSource = Recclubmem1
                    
            dcstu2.ListField = "StuID"
            Set dcstu2.RowSource = Recclubmem1
            
            
            If Not (Recclubmem1.EOF And Recclubmem1.BOF) Then
               Recclubmem1.MoveFirst
               lvclubmem.ListItems.clear
            
                While Not Recclubmem1.EOF
                    Set List = lvclubmem.ListItems.add
                    List.Text = Recclubmem1!stuid
                    List.SubItems(1) = Recclubmem1!FatherName + " " + Recclubmem1!StudentName
                    List.SubItems(2) = Recclubmem1!JoinDate
                    List.SubItems(3) = Recclubmem1!Description
                    Recclubmem1.MoveNext
                Wend
            Else
                lvclubmem.ListItems.clear
            End If

            
            If Not (Recclubmain1.EOF And Recclubmain1.BOF) Then
            Recclubmain1.MoveFirst
            lvmain.ListItems.clear
            While Not Recclubmain1.EOF
            Set List = lvmain.ListItems.add
            
            List.Text = Recclubmain1!Years
            List.SubItems(1) = Recclubmain1!Super_StaffID
            List.SubItems(2) = Recclubmain1!Super_StafName
            List.SubItems(3) = Recclubmain1!Pres_StuID
            List.SubItems(4) = Recclubmain1!Pres_StuName
            List.SubItems(5) = Recclubmain1!Sec_StuID
            List.SubItems(6) = Recclubmain1!Sec_StuName
            List.SubItems(7) = Recclubmain1!DESCRIPTIONS
            
            Recclubmain1.MoveNext
            Wend
            Else
            lvmain.ListItems.clear
            End If
        End If
    Else
    Frame6.Visible = False
    Command5.Enabled = False
    Command6.Enabled = False
    End If
End Sub

Private Sub dcStaff_Change()
On Error Resume Next
Dim stfname As String
   
    stfname = Trim(dcStaff.Text)
    
           If stfname <> "" Then
                RecStaff.MoveFirst
                RecStaff.Find "StaffID = '" & stfname & "'"
                If RecStaff.EOF Then
                    Text2.Text = ""
                Else
                    Text2.Text = RecStaff!FullName
                    
               End If
               Else
               Text2.Text = ""
            End If
End Sub

Private Sub dcstu1_Change()
On Error Resume Next
Dim stuid As String
   
    stuid = Trim(dcstu1.Text)
    
           If stuid <> "" Then
                RecStu.MoveFirst
                RecStu.Find "StuID = '" & stuid & "'"
                If RecStu.EOF Then
                    Text3.Text = ""
                Else
                    Text3.Text = RecStu!FatherName + " " + RecStu!StudentName
                    
               End If
               Else
               Text3.Text = ""
            End If
End Sub

Private Sub dcstu2_Change()
On Error Resume Next
Dim stuid As String
   
    stuid = Trim(dcstu2.Text)
    
           If stuid <> "" Then
                RecStu.MoveFirst
                RecStu.Find "StuID = '" & stuid & "'"
                If RecStu.EOF Then
                    Text4.Text = ""
                Else
                    Text4.Text = RecStu!FatherName + " " + RecStu!StudentName
                    
               End If
               Else
               Text4.Text = ""
            End If
End Sub

Private Sub Form_Load()

On Error Resume Next
Call modform.FormSize(Me, "Maintance CLUBS/UNIONS")

Set RecStaff = openDB.OpenRecord("SELECT * FROM STAFF")
'Set RecStu = openDB.OpenRecord("SELECT A.StuID,M.StudentName,M.FatherName,A.Curr_Class FROM MAINSTUDENTS M,ACTIVESTUDENT A where M.TemID=A.TemID and Old_Status='No'")
Set RecStu = openDB.OpenRecord("select A.StuID,M.StudentName,M.FatherName,A.Curr_Class  from MAINSTUDENTS M,ACTIVESTUDENT A where M.StuID=A.StuID")
Set Recclub = openDB.OpenRecord("SELECT * FROM CLUB")
Set Recclubmain = openDB.OpenRecord("SELECT * FROM CLUBMAINTAINCE")
Set Recclubmem = openDB.OpenRecord("select * from CLUBMEMBER")

dcclubname.ListField = "CName"
Set dcclubname.RowSource = Recclub

dcStaff.ListField = "StaffID"
Set dcStaff.RowSource = RecStaff

dcclubmem.ListField = "StuID"
Set dcclubmem.RowSource = RecStu

Command5.Enabled = False
Command6.Enabled = False

Frame7.Visible = False
Frame6.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

RecStu.Close
RecStaff.Close
Recclub.Close
Recclubmain.Close
Recclubmem.Close
Recclubmem1.Close
Recclubmain1.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

