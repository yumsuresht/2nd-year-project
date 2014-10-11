VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmlibpayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   Icon            =   "frmlibpayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7920
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Payment"
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   6015
      Begin VB.OptionButton Option3 
         Caption         =   "Others"
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fine"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin MSComctlLib.ListView lvPayment1 
         Height          =   3855
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   6800
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MemID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Member Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Post"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Frame Frame7 
         Caption         =   "Operations"
         Height          =   2175
         Left            =   6120
         TabIndex        =   24
         Top             =   0
         Width           =   1575
         Begin VB.CommandButton Command3 
            Caption         =   "Ok"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1335
         Left            =   0
         TabIndex        =   12
         Top             =   840
         Width           =   6015
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4200
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   720
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmlibpayment.frx":030A
            Left            =   4200
            List            =   "frmlibpayment.frx":0314
            TabIndex        =   13
            Top             =   720
            Width           =   1695
         End
         Begin MSDataListLib.DataCombo dcmemid1 
            Height          =   315
            Left            =   1440
            TabIndex        =   16
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label4 
            Caption         =   "Fine"
            Height          =   255
            Left            =   3240
            TabIndex        =   20
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Member ID"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Description"
            Height          =   255
            Left            =   3240
            TabIndex        =   17
            Top             =   840
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame6 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   0
         TabIndex        =   5
         Top             =   840
         Width           =   6015
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   720
            Width           =   4455
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dcmemid 
            Height          =   315
            Left            =   1440
            TabIndex        =   8
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label1 
            Caption         =   "Member ID"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Fine"
            Height          =   255
            Left            =   4080
            TabIndex        =   9
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Operations"
         Height          =   2175
         Left            =   6120
         TabIndex        =   3
         Top             =   0
         Width           =   1575
         Begin VB.CommandButton Command2 
            Caption         =   "Ok"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView lvPayment 
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   6800
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
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MemberID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Member Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "No of days"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fine (Rs)"
            Object.Width           =   1940
         EndProperty
      End
   End
End
Attribute VB_Name = "frmlibpayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecRecep As ADODB.Recordset
Dim RecMem As ADODB.Recordset
Dim RecLend As ADODB.Recordset
Dim RecID As ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If (dcmemid.Text = "" Or Text1.Text = "") Then
MsgBox "Invalid Member ID"
Exit Sub
End If

RecLend!FinePaid = "PAID"
RecLend!LendStatus = "YES"
RecLend.UpdateBatch
RecLend.Requery
dbglend.Refresh

RecRecep.AddNew
RecRecep!RecNo = RecID!RECEPID + 1
RecRecep!memid = dcmemid.Text
RecRecep!Payment = Val(Text2.Text)
RecRecep!Pay_Status = "Fine Fee "
RecRecep.UpdateBatch
RecRecep.Requery


RecID!RECEPID = Val(RecID!RECEPID + 1)
RecID.UpdateBatch
RecID.Requery
Call filllist
End Sub

Private Sub Command3_Click()
RecRecep.AddNew
RecRecep!RecNo = RecID!RECEPID + 1
RecRecep!memid = dcmemid1.Text
RecRecep!Payment = Val(Text3.Text)
RecRecep!Pay_Status = Combo1.Text + "Fine Fee "
RecRecep.UpdateBatch
RecRecep.Requery


RecID!RECEPID = Val(RecID!RECEPID + 1)
RecID.UpdateBatch
RecID.Requery
Call filllist1
End Sub

Private Sub dcmemid_Change()
Dim memid As String
On Error Resume Next
    memid = Trim(dcmemid.Text)
    If memid <> "" Then
        RecLend.MoveFirst
        RecLend.Find "MemID = '" & memid & "'"
        If RecLend.EOF Then
            Text1.Text = ""
            Text2.Text = ""
        Else
            Text1.Text = RecLend!MemName
            Text2.Text = RecLend!Fine
        End If
    Else
           Text1.Text = ""
           Text2.Text = ""
    
    End If
End Sub

Private Sub dcmemid1_Change()
Dim memid1 As String
On Error Resume Next
    memid1 = Trim(dcmemid1.Text)
    If memid1 <> "" Then
        RecMem.MoveFirst
        RecMem.Find "MemID = '" & memid1 & "'"
        If RecMem.EOF Then
            Text4.Text = ""
        Else
            Text4.Text = RecMem!MemName
        End If
    Else
           Text4.Text = ""
           
    
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "Library Late Payments")
    Set RecMem = openDB.OpenRecord("SELECT * FROM LIBRARYMEMBER")
    Set RecRecep = openDB.OpenRecord("SELECT * FROM PAYMENTS")
    Set RecLend = openDB.OpenRecord("select * from LENDING L,LIBRARYMEMBER M where M.MemID=L.MemID and FinePaid='NO'")
    Set RecID = openDB.OpenRecord("SELECT * FROM IDS")

    dcmemid.ListField = "MemID"
    Set dcmemid.RowSource = RecLend
        
    dcmemid1.ListField = "MemID"
    Set dcmemid1.RowSource = RecMem
    
    Call filllist
    Option1_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecRecep.Close
RecMem.Close
RecLend.Close
RecID.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Public Sub filllist()
On Error Resume Next


RecLend.MoveFirst
If Not (RecLend.EOF And RecLend.BOF) Then
    RecLend.MoveFirst
        lvPayment.ListItems.clear
        While Not RecLend.EOF
            Set List = lvPayment.ListItems.Add
            List.Text = RecLend.Fields(1)
            List.SubItems(1) = RecLend.Fields(11)
            If (Val(DateDiff("d", CDate(RecLend!DueDate), CDate(RecLend!ReturnDate))) > 0) Then
                List.SubItems(2) = DateDiff("d", CDate(RecLend!DueDate), CDate(RecLend!ReturnDate))
            Else
                List.SubItems(2) = " "
            End If
            List.SubItems(3) = RecLend.Fields(6)
            RecLend.MoveNext
        Wend
        Else
        lvPayment.ListItems.clear
    End If
End Sub

Private Sub lvPayment_Click()
On Error Resume Next
dcmemid.Text = lvPayment.SelectedItem.Text
If (Err.Number = 91) Then
MsgBox "No record found"
End If

End Sub

Private Sub lvPayment1_Click()
On Error Resume Next
dcmemid1.Text = lvPayment1.SelectedItem.Text
If (Err.Number = 91) Then
MsgBox "No record found"
End If

End Sub

Private Sub Option1_Click()
Frame6.Visible = True
Frame3.Visible = False
Call filllist1
End Sub

Private Sub Option3_Click()
Frame3.Visible = True
Frame6.Visible = False
End Sub

Public Sub filllist1()
On Error Resume Next
RecMem.MoveFirst
If Not (RecMem.EOF And RecMem.BOF) Then
    RecMem.MoveFirst
        lvPayment1.ListItems.clear
        While Not RecMem.EOF
            Set List = lvPayment1.ListItems.Add
            List.Text = RecMem.Fields(0)
            List.SubItems(1) = RecMem.Fields(2)
            List.SubItems(2) = RecMem.Fields(3)
            RecMem.MoveNext
        Wend
        Else
        lvPayment1.ListItems.clear
    End If
End Sub
