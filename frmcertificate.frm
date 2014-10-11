VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmcertificate 
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12480
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmcertificate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1215
      Left            =   4680
      TabIndex        =   15
      Top             =   5400
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2143
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmcertificate.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13320
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Printer Settings"
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10200
      Width           =   1335
   End
   Begin MSComctlLib.ListView listclub 
      Height          =   2535
      Left            =   4800
      TabIndex        =   11
      Top             =   7440
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Post Held"
         Object.Width           =   1596
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Club/Union"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Year"
         Object.Width           =   5292
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   6720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmcertificate.frx":04A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label add 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   11775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   10695
      Left            =   120
      Top             =   120
      Width           =   12255
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CONDUCT AND CHARACTER:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "POST HELD IN SCHOOL :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   2880
      Width           =   7095
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "CO - CURRICULAR ACTIVITIES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Caption         =   "EDUCATIONAL QUALIFICATION :"
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
      TabIndex        =   4
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "PERIOD OF EDUCATION"
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
      TabIndex        =   3
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "NAME :"
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
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "CERTIFICATE OF CHARACTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2160
      Width           =   5400
   End
   Begin VB.Label name1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   11655
   End
End
Attribute VB_Name = "frmcertificate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStu As ADODB.Recordset
Dim RecOlRes As ADODB.Recordset
Dim RecAlRes As ADODB.Recordset
Dim Recclubmem As ADODB.Recordset
Dim RecSchool As ADODB.Recordset

Dim STUID2 As String

Private Sub Command1_Click()
Command1.Visible = False
Command2.Visible = False
frmcertificate.PrintForm

End Sub

Private Sub Command2_Click()
CommonDialog1.ShowPrinter
End Sub



Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "Character Certificate")

Dim qual As String
STUID2 = frmCharacter.STUID1
'Set RecStu = openDB.OpenRecord("select M.StuID,M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin,O.D_Of_Leave,O.LastGrade from MAINSTUDENTS M, OLDBOYS O where M.StuID=O.StuID and M.StuID='" & STUID2 & "'")

Set RecStu = openDB.OpenRecord("select distinct(M.StuID) ,M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin AS Admission_Date,O.LastGrade,O.D_Of_Leave from MAINSTUDENTS M, OLDBOYS O where M.StuID=O.StuID and M.StuID='" & STUID2 & "' Union select distinct(M.StuID),M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin AS Admission_Date,A.Curr_Class,getdate()  from MAINSTUDENTS M, ACTIVESTUDENT A where M.StuID=A.StuID and M.StuID='" & STUID2 & "'")


Set RecOlRes = openDB.OpenRecord("select * from OLRESULT WHERE StuID='" & STUID2 & "'")
Set RecAlRes = openDB.OpenRecord("select * from ALRESULT WHERE StuID='" & STUID2 & "'")
Set Recclubmem = openDB.OpenRecord("select * from CLUBMEMBER where STUID='" & STUID2 & "'")
    Set RecSchool = openDB.OpenRecord("select * from SCHOOL")

name1.Caption = RecSchool!SCHOOLNAME
add.Caption = RecSchool!ADDRESS1 + " ," + RecSchool!ADDRESS2


Call fillclubdetails1
If (RecAlRes.RecordCount > 0) Then
Label10.Caption = "HELD THE G.C.E(A/L) EXAMINATION IN " & RecAlRes.Fields(3)
End If
If (RecOlRes.RecordCount > 0) Then
Label9.Caption = "HELD THE G.C.E(O/L) EXAMINATION IN " & RecOlRes.Fields(2)
Else
Label9.Caption = "STUDIED UNTILL GRADE " & RecStu.Fields(6)
End If
If (Recclubmem.RecordCount = 0) Then
Label11.Visible = False
Else
Label11.Visible = True
End If


Label7.Caption = UCase(RecStu.Fields(2) & " " & RecStu.Fields(1))
Label8.Caption = Year(RecStu.Fields(5)) & " TO " & Year(RecStu.Fields(7))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecStu.Close
RecOlRes.Close
RecAlRes.Close
Recclubmem.Close
RecSchool.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Public Sub fillclubdetails1()
On Error Resume Next
If Not (Recclubmem.EOF And Recclubmem.BOF) Then
    Recclubmem.MoveFirst
        listclub.ListItems.clear
        While Not Recclubmem.EOF
            Set List = listclub.ListItems.add
            List.Text = "Member"
            List.SubItems(1) = Recclubmem.Fields(0)
            List.SubItems(2) = " (" & Format(Recclubmem.Fields(2), "dd-mm-yyyy") & " to " & Format(Date, "dd-mm-yyyy") & " )"
            Recclubmem.MoveNext
        Wend
        Else
        listclub.ListItems.clear
    End If

End Sub

