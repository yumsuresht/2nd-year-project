VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmbooksearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   2040
   ClientLeft      =   3645
   ClientTop       =   4035
   ClientWidth     =   5505
   Icon            =   "frmbooksearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleMode       =   0  'User
   ScaleTop        =   3600
   ScaleWidth      =   5505
   Begin MSDataListLib.DataCombo dcbook 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   600
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4095
      Begin VB.OptionButton Option5 
         Caption         =   "All"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Category"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Title"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Book ID"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Access No"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Find What:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmbooksearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecBook, RecBook1, Recbook3 As ADODB.Recordset
Public sql As String

Private Sub Command1_Click()
On Error Resume Next

If (dcbook.Text = "") Then
MsgBox "Blank Field"
Exit Sub
End If
If (Option1.Value = True) Then
sql = "and C.AccessNo='" + dcbook.Text + "'"
ElseIf (Option2.Value = True) Then
sql = "and C.BookID=" + dcbook.Text
ElseIf (Option3.Value = True) Then
sql = "and B.Title  Like '%" + dcbook.Text + "%'"
ElseIf (Option4.Value = True) Then
sql = "and B.Catagory='" + dcbook.Text + "'"
ElseIf (Option5.Value = True) Then
sql = ""
End If
Set Recbook2 = openDB.OpenRecord("select B.BookID,C.AccessNo,B.Title,B.Edition,B.Catagory AS Category,C.BookStatus from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID " + sql)

    Recbook2.MoveFirst
    Set frmlend.dgbook1.DataSource = Recbook2
    frmlend.dgbook1.Caption = "Book Details"
    frmlend.dgbook1.HeadFont.Bold = True
    frmlend.dgbook1.HeadFont.Size = 10
    frmlend.dgbook1.Columns(0).Alignment = dbgCenter
   


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmlend.Enabled = False
Set RecBook = openDB.OpenRecord("select * from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID")
Set RecBook1 = openDB.OpenRecord("select Distinct(B.BookID) from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID")
Set Recbook3 = openDB.OpenRecord("select Distinct(B.Catagory) from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID")



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecBook.Close
RecBook1.Close
Recbook3.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmlend.Enabled = True
End Sub

Private Sub Option1_Click()
Set dcbook.RowSource = RecBook
dcbook.ListField = "AccessNo"
dcbook.Text = ""
End Sub

Private Sub Option2_Click()
Set dcbook.RowSource = RecBook1
dcbook.ListField = "BookID"
dcbook.Text = ""
End Sub

Private Sub Option3_Click()
dcbook.Text = "Enter the Book title"

End Sub

Private Sub Option4_Click()
Set dcbook.RowSource = Recbook3
dcbook.ListField = "Catagory"
dcbook.Text = ""

End Sub

Private Sub Option5_Click()
dcbook.Text = "ALL"

End Sub
