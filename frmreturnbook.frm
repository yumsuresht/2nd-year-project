VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmreturnbook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10545
   Icon            =   "frmreturnbook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10545
   Begin VB.Frame Frame4 
      Height          =   2775
      Left            =   8880
      TabIndex        =   19
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ok"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8655
      Begin VB.Frame Frame5 
         Caption         =   "Fines"
         Height          =   855
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   4935
         Begin VB.TextBox txtlost 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3360
            TabIndex        =   27
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtdam 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1200
            TabIndex        =   26
            Text            =   "0"
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Lost"
            Height          =   255
            Left            =   2520
            TabIndex        =   25
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Damage"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
      End
      Begin MSDataListLib.DataCombo dcmem 
         Height          =   315
         Left            =   1800
         TabIndex        =   22
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   5160
         TabIndex        =   3
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   290
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "Fine (Rs)"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
      End
      Begin MSDataListLib.DataCombo dcbook 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Access No"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Borrowed Date"
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Due Date"
         Height          =   255
         Left            =   5040
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Return Date"
         Height          =   255
         Left            =   5040
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "MemberID"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Member Name"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   10335
      Begin MSComctlLib.ListView lvReturn 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6376
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccessNo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "MemberID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Member Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Borrow Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Due Date"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmreturnbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecLend As ADODB.Recordset
Dim Reclend1 As ADODB.Recordset
Dim RecRecep As ADODB.Recordset
Dim RecID As ADODB.Recordset
Dim RecBook As ADODB.Recordset
Dim RecBook1 As ADODB.Recordset
Dim total As Integer




Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
frmlibpayment.Show
End Sub

Private Sub Command3_Click()
On Error Resume Next


If (Val(RecStuID!RECEPID) = 0) Then
MsgBox "You must initilized the Receipt No"
frminitial.Show
frminitial.Text5.SetFocus
frminitial.Text5.BackColor = &H80000018
Unload Me
Exit Sub
End If


Dim bookfine As Integer
Dim msg, cat As String
If (Option2.Value = True) Then
total = Val(txtdam.Text)
cat = "(Damage)"
End If
If (Option3.Value = True) Then
total = Val(txtlost.Text)
cat = "(Lost)"

End If


If (dcmem.Text = "" Or Text6.Text = "") Then
MsgBox "Balnk Fields"
Exit Sub
End If

If (Frame1.Visible = True And Val(Text7.Text) > 0 Or Val(txtdam.Text) > 0 Or Option3.Value = True) Then
msg = MsgBox("You want to pay Fine Rs(" & total + Val(Text7.Text) & "). Can you pay now?", vbYesNo)
    If (msg = vbYes) Then
    RecLend!ReturnDate = Date
    RecLend!Fine = total + Val(Text7.Text)
    RecLend!FinePaid = "PAID"
    RecLend!Transact = "COMPLETE"
    RecLend!BookStatus = "YES"
    RecLend!LendStatus = "YES"
    RecLend.UpdateBatch
    RecLend.Requery
    Reclend1.Requery
    dbglend.Refresh
    
    RecRecep.AddNew
    RecRecep!RecNo = Val(RecID!RECEPID) + 1
    RecRecep!memid = dcmem.Text
    RecRecep!Payment = total + Val(Text7.Text)
    RecRecep!Pay_Status = "Fine Fee " & cat
    RecRecep.UpdateBatch
    RecRecep.Requery
    
    Else
    
    RecLend!ReturnDate = Date
    RecLend!Fine = total + Val(Text7.Text)
    RecLend!FinePaid = "NO"
    RecLend!Transact = "COMPLETE"
    RecLend!BookStatus = "YES"
    RecLend!LendStatus = "FINE"
    RecLend.UpdateBatch
    RecLend.Requery
    Reclend1.Requery
    dbglend.Refresh
    End If
    If (Option3.Value = True) Then
        Call lostBook
    End If

Else
MsgBox "Ok"

    RecLend!ReturnDate = Date
    RecLend!Fine = 0
    RecLend!FinePaid = "PAID"
    RecLend!Transact = "COMPLETE"
    RecLend!BookStatus = "YES"
    RecLend!LendStatus = "YES"
    RecLend.UpdateBatch
    RecLend.Requery
    Reclend1.Requery
    dbglend.Refresh
End If

Call lostBook
modform.ClearTextBoxes (Me)
dcbook.Text = ""

Call filllist
End Sub

Private Sub Command4_Click()
End Sub

Private Sub dcbook_Change()
Dim bookno As String
Dim date1 As Date
Dim dateval As String
    On Error Resume Next
    bookno = Trim(dcbook.Text)
    
               If bookno <> "" Then
                RecLend.MoveFirst
                RecLend.Find "AccessNo = '" & bookno & "'"
                If RecLend.EOF Then
                    Text1.Text = ""
                    dcmem.Text = ""
                    dcmem.Text = ""
                    Text3.Text = ""
                    Text4.Text = ""
                    Text5.Text = ""
                    Frame1.Visible = False
                  
                Else
                    Text1.Text = RecLend!title
                    dcmem.Text = RecLend!memid
                    dcmem.Text = RecLend!memid
                    Text3.Text = RecLend!MemName
                    Text4.Text = Format(RecLend!BorrowDate, "dd/mm/yyyy")
                    Text5.Text = Format(RecLend!DueDate, "dd/mm/yyyy")
                    dateval = DateDiff("d", CDate(RecLend!DueDate), Date)


                    If (Val(dateval) > 0 And RecLend!status <> "Staff") Then
                        Frame1.Visible = True
                        Text7.Text = Val(dateval) * 5 + total
                    Else
                        Frame1.Visible = False
                    End If
                    
                End If
                Else
                    Text1.Text = ""
                    dcmem.Text = ""
                    dcmem.Text = ""
                    Text3.Text = ""
                    Text4.Text = ""
                    Text5.Text = ""
            End If
            
    Set RecBook = openDB.OpenRecord("SELECT * FROM COPY_OF_BOOK where AccessNo='" & Trim(dcbook.Text) & "'")
    Set RecBook1 = openDB.OpenRecord("select * from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID and C.AccessNo='" & Trim(dcbook.Text) & "'")
    txtlost.Text = RecBook1!Price
    Text2.Text = RecBook1!N_Of_Co
End Sub

Private Sub dcmem_Change()
Dim mem As String
On Error Resume Next
mem = Trim(dcmem.Text)
    
               If mem <> "" Then
                RecLend.MoveFirst
                RecLend.Find "MemID = '" & mem & "'"
                If RecLend.EOF Then
                    dcbook.Text = ""
                    Frame1.Visible = False
                  
                Else
                    dcbook.Text = RecLend!AccessNo
                                       
                End If
                Else
                    dcbook.Text = ""
                    Text1.Text = ""
                    dcmem.Text = ""
                    Text3.Text = ""
                    Text4.Text = ""
                    Text5.Text = ""
            End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "Return Book")
    Set Reclend1 = openDB.OpenRecord("select L.AccessNo,B.Title,L.MemID,M.MemName,L.BorrowDate,L.DueDate from LENDING L,COPY_OF_BOOK C,Book B ,LIBRARYMEMBER M Where L.AccessNo = C.AccessNo and B.BookID=C.BookID and M.MemID= L.MemID and C.BookStatus='LEND' and M.LendStatus='LEND' and L.Transact='LEND'")
    Set RecLend = openDB.OpenRecord("select * from LENDING L,COPY_OF_BOOK C,Book B ,LIBRARYMEMBER M Where L.AccessNo = C.AccessNo and B.BookID=C.BookID and M.MemID= L.MemID and C.BookStatus='LEND' and M.LendStatus='LEND' and L.Transact='LEND'")
    Set RecRecep = openDB.OpenRecord("select * from PAYMENTS")
    Set RecID = openDB.OpenRecord("SELECT * FROM IDS")
    
    dcbook.ListField = "AccessNo"
    Set dcbook.RowSource = RecLend
    
    
    dcmem.ListField = "MemID"
    Set dcmem.RowSource = RecLend
    
    Text6.Text = Format(Date, "dd/mm/yyyy")
    Call filllist
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecLend.Close
Reclend1.Close
RecRecep.Close
RecID.Close
RecBook1.Close
RecBook.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub


Public Sub filllist()
On Error Resume Next
Reclend1.MoveFirst
If Not (Reclend1.EOF And Reclend1.BOF) Then
    Reclend1.MoveFirst
        lvReturn.ListItems.clear
        While Not Reclend1.EOF
            Set List = lvReturn.ListItems.Add
            List.Text = Reclend1.Fields(0)
            List.SubItems(1) = Reclend1.Fields(1)
            List.SubItems(2) = Reclend1.Fields(2)
            List.SubItems(3) = Reclend1.Fields(3)
            List.SubItems(4) = Format(Reclend1.Fields(4), "dd/mm/yyyy")
            List.SubItems(5) = Format(Reclend1.Fields(5), "dd/mm/yyyy")
            Reclend1.MoveNext
        Wend
        Else
        lvReturn.ListItems.clear
    End If
End Sub

Private Sub lvReturn_Click()
On Error Resume Next
dcbook.Text = lvReturn.SelectedItem.Text
dcmem.Text = lvReturn.SelectedItem.SubItems(2)

If (Err.Number = 91) Then
MsgBox "No record found"
End If

End Sub

Private Sub Option2_Click()
txtdam.Visible = True
txtlost.Visible = False
txtlost.Text = 0
End Sub

Private Sub Option3_Click()
txtdam.Visible = False
txtlost.Visible = True
txtdam.Text = 0
Call dcbook_Change
End Sub

Public Function lostBook()
On Error Resume Next
RecBook.Delete
RecBook.UpdateBatch
RecBook.Requery

RecBook1!N_Of_Co = Val(RecBook1!N_Of_Co) - 1
RecBook1.UpdateBatch
RecBook1.Requery
End Function
