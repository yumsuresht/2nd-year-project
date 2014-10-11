VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmaddbook 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Book"
   ClientHeight    =   8205
   ClientLeft      =   240
   ClientTop       =   1935
   ClientWidth     =   8700
   ClipControls    =   0   'False
   Icon            =   "frmaddbook1.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleMode       =   0  'User
   ScaleWidth      =   6170.214
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Book"
      TabPicture(0)   =   "frmaddbook1.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text14"
      Tab(0).Control(1)=   "cmddelete"
      Tab(0).Control(2)=   "Check1"
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(5)=   "Text3"
      Tab(0).Control(6)=   "Text5"
      Tab(0).Control(7)=   "Text6"
      Tab(0).Control(8)=   "Text7"
      Tab(0).Control(9)=   "Text8"
      Tab(0).Control(10)=   "Text9"
      Tab(0).Control(11)=   "Text4"
      Tab(0).Control(12)=   "Command2"
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(14)=   "cmdadd"
      Tab(0).Control(15)=   "clr"
      Tab(0).Control(16)=   "exit"
      Tab(0).Control(17)=   "edit"
      Tab(0).Control(18)=   "Command3"
      Tab(0).Control(19)=   "Frame1"
      Tab(0).Control(20)=   "dacom1"
      Tab(0).Control(21)=   "dacom"
      Tab(0).Control(22)=   "Label6"
      Tab(0).Control(23)=   "Label22"
      Tab(0).Control(24)=   "Label21"
      Tab(0).Control(25)=   "Label20"
      Tab(0).Control(26)=   "Label19"
      Tab(0).Control(27)=   "Label18"
      Tab(0).Control(28)=   "Label17"
      Tab(0).Control(29)=   "Label16"
      Tab(0).Control(30)=   "Label15"
      Tab(0).Control(31)=   "Label14"
      Tab(0).Control(32)=   "Label13"
      Tab(0).Control(33)=   "Label12"
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Copy of Book"
      TabPicture(1)   =   "frmaddbook1.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "dcbook"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text10"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command6"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "dgcopybook"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command7"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame3"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Text13"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71640
         TabIndex        =   12
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         TabIndex        =   23
         Top             =   600
         Width           =   615
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   240
         TabIndex        =   49
         Top             =   1800
         Width           =   2415
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "New Copies"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   2760
         TabIndex        =   47
         Top             =   1800
         Width           =   3015
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1080
            TabIndex        =   22
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Access No"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Search"
         Height          =   375
         Left            =   7080
         TabIndex        =   26
         Top             =   1560
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid dgcopybook 
         Height          =   5055
         Left            =   240
         TabIndex        =   46
         Top             =   2760
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   8916
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
      Begin VB.CommandButton Command6 
         Caption         =   "Delete"
         Height          =   375
         Left            =   7080
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7080
         TabIndex        =   27
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "New"
         Height          =   375
         Left            =   7080
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   20
         Top             =   1080
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo dcbook 
         Height          =   315
         Left            =   1440
         TabIndex        =   19
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete"
         Height          =   375
         Left            =   -68160
         TabIndex        =   14
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Over Night"
         Height          =   255
         Left            =   -70080
         TabIndex        =   32
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74880
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71880
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74880
         TabIndex        =   3
         Top             =   1200
         Width           =   6375
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74880
         TabIndex        =   4
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   6
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72720
         TabIndex        =   7
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70680
         TabIndex        =   8
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74880
         TabIndex        =   11
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70680
         TabIndex        =   5
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000013&
         Caption         =   ">>"
         Height          =   375
         Left            =   -67320
         TabIndex        =   31
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "<<"
         Height          =   375
         Left            =   -68160
         TabIndex        =   30
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton cmdadd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add"
         Height          =   375
         Left            =   -68160
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton clr 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear"
         Height          =   375
         Left            =   -68160
         TabIndex        =   17
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton exit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   -68160
         TabIndex        =   18
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edit"
         Height          =   375
         Left            =   -68160
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   375
         Left            =   -68160
         TabIndex        =   16
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   28
         Top             =   3960
         Width           =   8295
         Begin MSDataGridLib.DataGrid dtgCode 
            Height          =   3615
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   6376
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
      Begin MSDataListLib.DataCombo dacom1 
         Height          =   315
         Left            =   -71640
         TabIndex        =   10
         Top             =   3000
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dacom 
         Height          =   315
         Left            =   -74880
         TabIndex        =   9
         Top             =   3000
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "Price"
         Height          =   255
         Left            =   -71640
         TabIndex        =   52
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Available Copies"
         Height          =   255
         Left            =   3840
         TabIndex        =   51
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "BookName"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "BookID"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Book ID"
         Height          =   255
         Left            =   -74880
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "ISBN"
         Height          =   255
         Left            =   -71880
         TabIndex        =   42
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Publisher"
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Year of Publish"
         Height          =   255
         Left            =   -74880
         TabIndex        =   40
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Book Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   39
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Author"
         Height          =   255
         Left            =   -70680
         TabIndex        =   38
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Category"
         Height          =   255
         Left            =   -71640
         TabIndex        =   37
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Number of Copies"
         Height          =   255
         Left            =   -70680
         TabIndex        =   36
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Edition"
         Height          =   255
         Left            =   -72720
         TabIndex        =   35
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Language"
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Donated By"
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   3360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmaddbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec, rec1, Rec2, Rec3, Rec4, Rec5, Rec6 As ADODB.Recordset
Dim msg As String



Private Sub clr_Click()
    Call modform.ClearTextBoxes(Me)
End Sub

Private Sub cmdadd_Click()
 On Error Resume Next
 Dim i As Integer
 rec.MoveFirst
 Rec2.MoveFirst
 
 If (Val(Text8.Text) = 0) Then
 MsgBox "Check your No of copies,You must enter at least 1 copy", vbInformation
 Exit Sub
 End If
 
 
 'Text1.Text = Rec2!BOOKIDS + 1
    rec.Find "BookID = '" & Trim(Text1.Text) & "'"
    If rec.EOF Then
        rec.AddNew
        'Rec!BOOKID = Rec2!BOOKIDS + 1
        rec!BookID = Text1.Text
        rec!ISBN = Text2.Text
        rec!title = Text3.Text
        rec!AuthorName = Text4.Text
        rec!publisher = Text5.Text
        rec!Y_OF_p = Val(Text6.Text)
        rec!edition = Val(Text7.Text)
        rec!Language = dacom.Text
        rec!Catagory = dacom1.Text
        rec!N_Of_Co = Val(Text8.Text)
        rec!Donation = Text9.Text
        rec!Overnight = Check1.Value
        rec!Price = Val(Text14.Text)
        rec.UpdateBatch
        rec1.Requery
        Rec4.Requery
        
         
        For i = 1 To Val(Text8.Text)
        Rec3.AddNew
        Rec3!BookID = Text1.Text
        Rec3!AccessNo = Text1.Text & "-" & i
        'Rec3!Status = "YES"
        Rec3.UpdateBatch
        Next i
        
        
        
        
        'Rec2!BOOKIDS = Rec2!BOOKIDS + 1
        'Rec2.UpdateBatch
      main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
          Call modform.ClearTextBoxes(Me)

      Else
        MsgBox "Duplicate ID found....!"
    End If
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
If MsgBox("Are you sure do you want to delete the current record?", vbYesNo, "Delete") = vbYes Then
        rec.Delete
        rec.UpdateBatch
        rec.Requery
        rec1.Requery
        Rec4.Requery
        display
    End If
    
    
End Sub

Private Sub Command1_Click()
On Error Resume Next
rec.MovePrevious
    If rec.BOF Then rec.MoveFirst
    Call display
End Sub

Private Sub Command2_Click()
On Error Resume Next
rec.MoveNext
    If rec.EOF Then rec.MoveLast
    Call display
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim co, i As Integer
Dim id As String
If (Command4.Caption = "New") Then
Command4.Caption = "Save"
Frame3.Visible = True
Else

Command4.Caption = "New"
Frame3.Visible = False
If (Text11.Text = "") Then
MsgBox "Check the New copy"
Exit Sub
End If
co = Val(rec!N_Of_Co) + Val(Text11.Text)
rec!N_Of_Co = co
rec.UpdateBatch

For i = 1 To co
    id = rec!BookID & "-" & i
        Rec5.MoveFirst
        Rec5.Find "AccessNo = '" + (Trim(id)) + "'"
        If Rec5.EOF Then
            Rec3.AddNew
            Rec3!BookID = rec!BookID
            Rec3!AccessNo = Trim(id)
            
            
        End If
Next i
Rec3.UpdateBatch
Rec3.Requery
dgcopybook.Refresh
 Call dcbook_Change
 Text11.Text = ""
End If

Frame2.Visible = False
Command6.Enabled = False

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
On Error Resume Next
If MsgBox("Are you sure do you want to delete the current record?", vbYesNo, "Delete") = vbYes Then
        Rec6!N_Of_Co = Rec6!N_Of_Co - 1
        Rec6.UpdateBatch
        Rec6.Requery
        Rec3.Delete
        Rec3.UpdateBatch
        dcbook.Text = ""
    End If
              
        rec.Requery
        rec1.Requery
        Rec2.Requery
        Rec3.Requery
        Rec4.Requery
        Rec5.Requery
        
        Command6.Enabled = False
        Frame2.Visible = False


End Sub

Private Sub Command7_Click()
Dim s As String
On Error Resume Next
s = InputBox("ENTER ACCESS NO", SEARCH)
If (s <> "") Then
 Rec3.MoveFirst
    Rec3.Find "AccessNo = '" + (Trim(s)) + "'"
   Set Rec6 = openDB.OpenRecord("select * from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID and C.AccessNo='" + Trim(s) + "'")
    If Rec3.EOF Then
        MsgBox "CAN NOT FIND"
        Else
        dcbook.Text = Rec6!title
        Text10.Text = Rec3!BookID
        Text12.Text = Rec3!AccessNo
        
        
        Rec6.MoveFirst
    Set dgcopybook.DataSource = Rec6
    dgcopybook.Caption = "Details of Books"
    dgcopybook.HeadFont.Bold = True
    dgcopybook.HeadFont.Size = 10
    dgcopybook.Columns(0).Alignment = dbgCenter
        
     Command6.Enabled = True
     Frame2.Visible = True
    End If
    
End If
Frame3.Visible = False
End Sub

Private Sub dcbook_Change()
On Error Resume Next
Dim title As String
title = Trim(dcbook.Text)
      
    If title <> "" Then
        rec.MoveFirst
                rec.Find "Title = '" & title & "'"
                If rec.EOF Then
                    Text10.Text = ""
                    Text13.Text = ""
                    Set Rec5 = openDB.OpenRecord("select B.BookID,C.AccessNo,B.Title,B.ISBN,B.Catagory AS Category from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID and B.Title= '" + Trim(dcbook.Text) + "'")
                    Rec5.MoveFirst
                    Set dgcopybook.DataSource = Rec5
                Else
                    Text10.Text = rec!BookID
                    Text13.Text = rec!N_Of_Co
                    Set Rec5 = openDB.OpenRecord("select B.BookID,C.AccessNo,B.Title,B.ISBN,B.Catagory AS Category from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID and B.Title= '" + Trim(dcbook.Text) + "'")
                    Rec5.MoveFirst
                    Set dgcopybook.DataSource = Rec5
                    
                End If
    End If

End Sub

Private Sub edit_Click()
 On Error GoTo Errhand
    rec!ISBN = Text2.Text
    rec!title = Text3.Text
    rec!AuthorName = Text4.Text
    rec!publisher = Text5.Text
    rec!Y_OF_p = Val(Text6.Text)
    rec!edition = Val(Text7.Text)
    rec!Language = dacom.Text
    rec!Catagory = dacom1.Text
    rec!N_Of_Co = Text8.Text
    rec!Donation = Text9.Text
    rec!Overnight = Check1.Value
    rec!Price = Val(Text14.Text)
    rec.UpdateBatch
    main.stbMain.Panels(1).Text = "Staus: Record successfully saved"
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub exit_Click()
Unload Me

End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "ADD BOOK")
    Set rec = openDB.OpenRecord("SELECT * FROM book")
    Set rec1 = openDB.OpenRecord("SELECT Distinct(Catagory) FROM book")
    Set Rec4 = openDB.OpenRecord("SELECT Distinct(Language) FROM book")
    Set Rec2 = openDB.OpenRecord("SELECT BOOKIDS FROM IDS")
    Set Rec3 = openDB.OpenRecord("SELECT * FROM COPY_OF_BOOK")


    dacom.ListField = "Language"
    Set dacom.RowSource = Rec4
    
    dacom1.ListField = "Catagory"
    Set dacom1.RowSource = rec1
    
    rec.MoveFirst
    dcbook.ListField = "Title"
    Set dcbook.RowSource = rec
    
   
    rec.MoveFirst
    Set dtgCode.DataSource = rec
    dtgCode.Caption = "Details of Books"
    dtgCode.HeadFont.Bold = True
    dtgCode.HeadFont.Size = 10
    dtgCode.Columns(0).Alignment = dbgCenter
   
    Command6.Enabled = False
    Frame2.Visible = False
    Frame3.Visible = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    main.stbMain.Panels(1).Text = "Status : Not Ready"

rec.Close
rec1.Close
Rec2.Close
Rec3.Close
Rec4.Close
Rec5.Close
Rec6.Close

End Sub

Public Sub display()
 On Error Resume Next

Text1.Text = rec!BookID
    Text2.Text = rec!ISBN
    Text3.Text = rec!title
    Text4.Text = rec!AuthorName
    Text5.Text = rec!publisher
    Text6.Text = rec!Y_OF_p
    Text7.Text = rec!edition
    dacom.Text = rec!Language
    dacom1.Text = rec!Catagory
    Text8.Text = rec!N_Of_Co
    Text9.Text = rec!Donation
    Check1.Value = rec!Overnight
    Text14.Text = rec!Price
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Text2.SetFocus
Else
    msg = MsgBox("You must enter number value", vbExclamation)
    KeyAscii = 0
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    Text7.SetFocus
Else
    msg = MsgBox("You must enter number value", vbExclamation)
    KeyAscii = 0
End If

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45 Then
ElseIf KeyAscii = 13 Then
    dacom.SetFocus
Else
    msg = MsgBox("You must enter number value", vbExclamation)
    KeyAscii = 0
End If
End Sub
