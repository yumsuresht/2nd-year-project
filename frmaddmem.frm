VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmaddmem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Member"
   ClientHeight    =   6720
   ClientLeft      =   240
   ClientTop       =   1935
   ClientWidth     =   6975
   ClipControls    =   0   'False
   Icon            =   "frmaddmem.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   6720
   ScaleMode       =   0  'User
   ScaleWidth      =   17856
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton frmexit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   20
      TabIndex        =   13
      Top             =   3000
      Width           =   6855
      Begin MSDataGridLib.DataGrid dtgCode 
         Height          =   2775
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4895
         _Version        =   393216
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "New Member"
      TabPicture(0)   =   "frmaddmem.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtname"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtmeid"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dcids"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   3495
         Begin VB.TextBox txtpay 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtcla 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Payment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Class"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSDataListLib.DataCombo dcids 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtmeid 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Student / Staff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   2775
         Begin VB.OptionButton Option1 
            Caption         =   "Student"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Staff"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Student ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Member ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmaddmem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rec As ADODB.Recordset
Public Rec1 As ADODB.Recordset
Public RecLibMem As ADODB.Recordset
Public RecLibMem1 As ADODB.Recordset
Public RecRecep As ADODB.Recordset
Public RecID As ADODB.Recordset


Private Sub cmdadd_Click()
On Error Resume Next

If (Val(RecStuID!MEMBERIDS) = 0) Then
MsgBox "You must initilized the MemberID"
frminitial.Show
frminitial.Text4.SetFocus
frminitial.Text4.BackColor = &H80000018
Unload Me
Exit Sub
End If



If (cmdadd.Caption = "Add") Then
cmdadd.Caption = "Update"
txtmeid.Text = RecID!MEMBERIDS + 1
ElseIf (cmdadd.Caption = "Update") Then
cmdadd.Caption = "Add"
If (dcids.Text = "" Or txtname.Text = "") Then
MsgBox "Invalid " & Label1.Caption
Exit Sub
End If

If (Option1.Value = True) Then
 RecLibMem.MoveFirst
 RecID.MoveFirst
    RecLibMem.Find "SCID = '" & Trim(dcids.Text) & "'"
    
    If RecLibMem.EOF Then
        RecLibMem.AddNew
        RecLibMem!memid = RecID!MEMBERIDS + 1
        RecLibMem!SCID = Trim(dcids.Text)
        RecLibMem!MemName = txtname.Text
        RecLibMem!status = "Student"
        RecLibMem!Membership = "YES"
        RecLibMem!StartClass = txtcla.Text
        
        RecLibMem.UpdateBatch
        RecLibMem.Requery
        RecLibMem1.Requery
        
        RecRecep.AddNew
        RecRecep!RecNo = RecID!RECEPID + 1
        RecRecep!memid = RecID!MEMBERIDS + 1
        RecRecep!Payment = Val(txtpay.Text)
        RecRecep!Pay_Status = "Member Fee " + " " + txtcla.Text
        RecRecep.UpdateBatch
        RecRecep.Requery
        
        RecID!MEMBERIDS = RecID!MEMBERIDS + 1
        RecID!RECEPID = RecID!RECEPID + 1
        RecID.UpdateBatch
        
        
        
        main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
        Call modform.ClearTextBoxes(Me)
      Else
        MsgBox "Allready added....!"
        Exit Sub
    End If
ElseIf (Option2.Value = True) Then
      RecID.MoveFirst
       RecLibMem.Find "SCID = '" & Trim(dcids.Text) & "'"
       If RecLibMem.EOF Then
           RecLibMem.AddNew
           RecLibMem!memid = RecID!MEMBERIDS + 1
           RecLibMem!SCID = Trim(dcids.Text)
           RecLibMem!MemName = txtname.Text
           RecLibMem!status = "Staff"
           RecLibMem!Membership = "YES"
           
           RecLibMem.UpdateBatch
           RecLibMem.Requery
           RecLibMem1.Requery
           RecID!MEMBERIDS = RecID!MEMBERIDS + 1
           RecID.UpdateBatch
           main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
           Call modform.ClearTextBoxes(Me)
         Else
           MsgBox "Allready added....!"
           Exit Sub
        End If
End If

End If

End Sub

Private Sub cmdclear_Click()
Call modform.ClearTextBoxes(Me)
cmdadd.Enabled = False
End Sub

Private Sub cmdex_Click()
Unload Me
End Sub

Private Sub dcids_Change()
Dim strCode As String
    On Error Resume Next
    strCode = Trim(dcids.Text)
    
    If (Option1.Value = True) Then
           If strCode <> "" Then
                rec.MoveFirst
                rec.Find "StuID = '" & strCode & "'"
                If rec.EOF Then
                    txtname.Text = ""
                    txtmeid.Text = ""
                    txtcla.Text = ""
                    txtpay.Text = ""
                   cmdadd.Enabled = False
                Else
                    txtname.Text = rec!StudentName + " " + rec!FatherName
                    txtcla.Text = rec!Curr_Class
                    s = Val(Left(txtcla.Text, 2))
                    If (s >= 6 And s <= 8) Then
                        txtpay.Text = 100
                        ElseIf (s > 8 And s <= 11) Then
                        txtpay.Text = 200
                    ElseIf (s > 11 And s <= 13) Then
                        txtpay.Text = 250
                    Else
                        txtpay.Text = 0
                    End If
                    cmdadd.Enabled = True
                    
               End If
            End If
    ElseIf (Option2.Value = True) Then
            If strCode <> "" Then
                Rec1.MoveFirst
                Rec1.Find "StaffID = '" & strCode & "'"
                If Rec1.EOF Then
                    txtname.Text = ""
                    txtmeid.Text = ""
                    Command6.Enabled = False
                    cmdadd.Enabled = False
                Else
                    txtname.Text = Rec1!FullName
                    Command6.Enabled = True
                    cmdadd.Enabled = True
               End If
            End If
    End If
    
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "Library Member")
    cmdadd.Enabled = False

    'Set Rec = openDB.OpenRecord("SELECT A.StuID,M.StudentName,M.FatherName,A.Curr_Class FROM MAINSTUDENTS M,ACTIVESTUDENT A where M.TemID=A.TemID and Old_Status='No'")
    Set rec = openDB.OpenRecord("select A.StuID,M.StudentName,M.FatherName,A.Curr_Class  from MAINSTUDENTS M,ACTIVESTUDENT A where M.StuID=A.StuID")
    
    Set Rec1 = openDB.OpenRecord("SELECT * FROM STAFF")
    Set RecLibMem = openDB.OpenRecord("SELECT * FROM LIBRARYMEMBER")
    Set RecLibMem1 = openDB.OpenRecord("SELECT MemID,SCID AS IDS,MemName As Member_Name,Status AS Designation FROM LIBRARYMEMBER")
    Set RecID = openDB.OpenRecord("SELECT * FROM IDS")
    Set RecRecep = openDB.OpenRecord("SELECT * FROM PAYMENTS")
    
    Option1_Click
    RecLibMem1.MoveFirst
    Set dtgCode.DataSource = RecLibMem1
    dtgCode.Caption = "Member Details"
    dtgCode.HeadFont.Bold = True
    dtgCode.HeadFont.Size = 10
    dtgCode.Columns(0).Alignment = dbgCenter
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    main.stbMain.Panels(1).Text = "Status : Not Ready"

rec.Close
Rec1.Close
RecLibMem.Close
RecLibMem1.Close
RecRecep.Close
RecID.Close
End Sub

Private Sub frmexit_Click()
Unload Me
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Label1.Caption = "Student ID"
Frame4.Visible = True
Label4.Visible = True
Label5.Visible = True

Set dcids.RowSource = rec
dcids.ListField = "StuID"
dcids.Text = ""

End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Label1.Caption = "Staff ID"
Frame4.Visible = False
Label4.Visible = False
Label5.Visible = False

Set dcids.RowSource = Rec1
dcids.ListField = "StaffID"
dcids.Text = ""

End If
End Sub

Private Sub Option4_Click()

End Sub

Private Sub txtstuid_KeyPress(KeyAscii As Integer)



End Sub

