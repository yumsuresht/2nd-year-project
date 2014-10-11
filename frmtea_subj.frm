VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmtea_subj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   Icon            =   "frmtea_subj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10425
   Begin VB.Frame Frame3 
      Height          =   3495
      Left            =   50
      TabIndex        =   17
      Top             =   2400
      Width           =   8895
      Begin MSComctlLib.ListView listteasub 
         Height          =   3135
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
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
            Text            =   "StaffID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SubjectID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Category"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Teachers"
      Height          =   2175
      Left            =   50
      TabIndex        =   10
      Top             =   120
      Width           =   5295
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   3135
      End
      Begin MSDataListLib.DataCombo dcstf 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "StaffID"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Description"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subject"
      Height          =   2175
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo dccategory 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcsub 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Category"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "ID"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   9120
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmtea_subj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset
Dim Rec1, Rec2, Rec3, Rec4 As ADODB.Recordset

Private Sub Command1_Click()
    On Error Resume Next

 rec.MoveFirst
    
        rec.AddNew
        rec!staffid = dcstf.Text
        rec!SubjectID = Text3.Text
        rec!Category = dccategory.Text
        rec!DESCRIPTIONS = Text2.Text
        rec.UpdateBatch
        rec.Requery
        Call filllist

  If (Err.Number = 3021 Or Err.Number = 0) Then
        main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
        Call modform.ClearTextBoxes(Me)
  Else
    MsgBox "duplicate"
  End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Command3_Click()
Dim s As String
On Error Resume Next
s = InputBox("ENTER FIND VALUE", SEARCH)
If (s <> "") Then
    rec.MoveFirst
    rec.Find "StaffID = '" & Trim(s) & "'"
    If rec.EOF Then
        MsgBox "CAN NOT FIND"
        dcstf.Text = ""
        Text3.Text = ""
        dccategory.Text = ""
        Text2.Text = ""
    Else
    If MsgBox("Are you sure do you want to delete the current record?", vbYesNo, "Delete") = vbYes Then
        rec.Delete
        rec.UpdateBatch
        rec.Requery
        Call filllist
    End If
        
    End If
Else
    dcstf.Text = ""
    Text3.Text = ""
    dccategory.Text = ""
    Text2.Text = ""
End If
End Sub

Private Sub dccategory_Change()
On Error Resume Next
Dim cate As String
cate = Trim(dccategory.Text)
      
    If cate <> "" Then
        Rec3.MoveFirst
                Rec3.Find "Category = '" & cate & "'"
                If Rec3.EOF Then
                    Text3.Text = ""
                    dcsub.Text = ""
                Else
                    Text3.Text = ""
                    dcsub.Text = ""
                    Rec2.Close
                    Set Rec2 = openDB.OpenRecord("SELECT * FROM SUBJECT where Category = '" & cate & "'")
                    dcsub.ListField = "SubjectNames"
                    Set dcsub.RowSource = Rec2
                    
                End If
    End If
End Sub

Private Sub dcstf_Change()
Dim strCode As String
    On Error Resume Next
    strCode = Trim(dcstf.Text)
    
           If strCode <> "" Then
                Rec1.MoveFirst
                Rec1.Find "StaffID = '" & strCode & "'"
                If Rec1.EOF Then
                    Text1.Text = ""
                Else
                    Text1.Text = Rec1!FullName
                    
               End If
            End If
    
End Sub

Private Sub dcsub_Change()

On Error Resume Next
Dim name, cate As String
name = Trim(dcsub.Text)
cate = Trim(dccategory.Text)
   If name <> "" Then
        Rec2.MoveFirst
                Rec2.Find "SubjectNames = '" & name & "'"
                If Rec2.EOF Then
                    Text3.Text = ""
                Else
                    
                    Set Rec4 = openDB.OpenRecord("SELECT * FROM SUBJECT where Category = '" & cate & "' and SubjectNames = '" & name & "'")
                    Text3.Text = Rec4!SubjectID
                    
                End If
    End If
End Sub



Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "Teacher-Subject")
    Set rec = openDB.OpenRecord("SELECT * FROM TEACHSUBJECT")
    Set Rec1 = openDB.OpenRecord("SELECT StaffID,FullName,Work_Hours FROM STAFF WHERE PostHeld='TEACHER'")
    Set Rec3 = openDB.OpenRecord("select DISTINCT(Category) from SUBJECT")
    
    Rec1.MoveFirst
    dcstf.ListField = "StaffID"
    Set dcstf.RowSource = Rec1
    
    Call filllist
    
    
    
    
    Rec3.MoveFirst
    dccategory.ListField = "Category"
    Set dccategory.RowSource = Rec3
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
rec.Close
Rec1.Close
Rec2.Close
Rec3.Close
Rec4.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"


End Sub


Public Sub filllist()
If Not (rec.EOF And rec.BOF) Then
    rec.MoveFirst
        listteasub.ListItems.clear
        While Not rec.EOF
            Set List = listteasub.ListItems.Add
            List.Text = rec.Fields(0)
            List.SubItems(1) = rec.Fields(1)
            List.SubItems(2) = rec.Fields(2)
            List.SubItems(3) = rec.Fields(3)
            rec.MoveNext
        Wend
        Else
        listteasub.ListItems.clear
    End If
End Sub
