VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmschsubject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   Icon            =   "frmschsubject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   10980
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   50
      TabIndex        =   7
      Top             =   4920
      Width           =   10815
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   9240
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin MSComctlLib.ListView listsubjects 
         Height          =   3855
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16776960
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subject ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Subject Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Category"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.ListBox listsub 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Frame Frame5 
      Caption         =   "Selected Subjects"
      Height          =   3795
      Left            =   6000
      TabIndex        =   2
      Top             =   960
      Width           =   4600
      Begin VB.ListBox dlsub 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "All Subjects"
      Height          =   3795
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4600
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   50
      TabIndex        =   0
      Top             =   120
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   5
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Core Subjects"
      TabPicture(0)   =   "frmschsubject.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Religions"
      TabPicture(1)   =   "frmschsubject.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Aesthetic Subjects"
      TabPicture(2)   =   "frmschsubject.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Additional Subjects"
      TabPicture(3)   =   "frmschsubject.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Technical Stream"
      TabPicture(4)   =   "frmschsubject.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Commerce Stream"
      TabPicture(5)   =   "frmschsubject.frx":0396
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Technical Subjects/Agriculture Stream"
      TabPicture(6)   =   "frmschsubject.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Home Economics Stream"
      TabPicture(7)   =   "frmschsubject.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
   End
End
Attribute VB_Name = "frmschsubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset
Dim rec1 As ADODB.Recordset
Dim Rec2 As ADODB.Recordset
Dim Rec3 As ADODB.Recordset
Dim Rec4 As ADODB.Recordset
Dim Rec5 As ADODB.Recordset


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If (MsgBox("Are you sure to select this subject?", vbOKCancel) = vbOK) Then
Dim name As String
name = listsub.Text

Set Rec3 = openDB.OpenRecord("select * from OLSUBJECT where Category= '" + Trim(SSTab1.Caption) + "' and SubjectNames='" + name + "'")

If listsub.ListIndex >= 0 Then
        dlsub.AddItem listsub.Text
        listsub.RemoveItem listsub.ListIndex
        
        
        rec1.MoveFirst
        rec1.AddNew
        rec1!SubjectID = Val(Rec2!SubjectID) + 1
        rec1!SubjectNames = Trim(name)
        rec1!Category = Trim(SSTab1.Caption)
        rec1.UpdateBatch
                
        Rec3!status = "NO"
        Rec3.UpdateBatch
                
        Rec2!SubjectID = Rec2!SubjectID + 1
        Rec2.UpdateBatch
        
        rec1.Requery
        Rec2.Requery
        Rec3.Requery
        Rec5.Requery
        rec.Requery
    End If
    
   Rec3.Close
       Call filllist

Else
Exit Sub
End If

End Sub

Private Sub Command3_Click()
On Error Resume Next
If (MsgBox("Are you sure to remove this subject?", vbOKCancel) = vbOK) Then

Dim name As String
name = dlsub.Text

Set Rec3 = openDB.OpenRecord("select * from OLSUBJECT where Category= '" + Trim(SSTab1.Caption) + "' and SubjectNames='" + name + "'")
Set Rec4 = openDB.OpenRecord("select * from SUBJECT where Category= '" + Trim(SSTab1.Caption) + "' and SubjectNames='" + name + "'")

        Rec3!status = "YES"
        Rec3.UpdateBatch
        
        Rec4.Delete
        Rec4.UpdateBatch


Dim i As Integer

    If dlsub.SelCount = 1 Then
        listsub.AddItem dlsub.Text
        dlsub.RemoveItem dlsub.ListIndex
    ElseIf dlsub.SelCount > 1 Then
        For i = dlsub.ListCount - 1 To 0 Step -1
            If dlsub.Selected(i) Then
                listsub.AddItem dlsub.List(i)
                dlsub.RemoveItem i
            End If
        Next
    End If
    
    rec1.Requery
    Rec2.Requery
    Rec3.Requery
    Rec4.Requery
    Rec5.Requery
    rec.Requery
       
    Rec3.Close
    Rec4.Close
    
    Call filllist
Else
Exit Sub
End If


End Sub

Private Sub dlsub_DblClick()
Command3_Click

End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "O/L Subjects")
SSTab1.Tab = 0
SSTab1_Click (0)
    Set Rec2 = openDB.OpenRecord("select * from IDS")

  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

rec.Close
rec1.Close
Rec2.Close
Rec3.Close
Rec4.Close
Rec5.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Public Sub selectsql(cat As String)
On Error Resume Next
rec.Close
rec1.Close
    Set rec = openDB.OpenRecord("select SubjectNames from OLSUBJECT where Status='YES' and Category='" + Trim(cat) + "'")
    Call Adddata(listsub, rec)
    Set rec1 = openDB.OpenRecord("select * from SUBJECT where Category='" + Trim(cat) + "'")
    Call Adddata(dlsub, rec1)
    Call filllist
   
End Sub

Public Sub Adddata(subname As ListBox, rec As ADODB.Recordset)
subname.clear
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        While Not rec.EOF
            subname.AddItem rec!SubjectNames

            rec.MoveNext
        Wend
    End If
End Sub

Private Sub listsub_DblClick()
Command2_Click

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If (SSTab1.Caption = "Additional Subjects") Then
    Call selectsql("Additional Subjects")
    
ElseIf (SSTab1.Caption = "Technical Stream") Then
    Call selectsql("Technical Stream")

ElseIf (SSTab1.Caption = "Commerce Stream") Then
    Call selectsql("Commerce Stream")
    
ElseIf (SSTab1.Caption = "Technical Subjects/Agriculture Stream") Then
    Call selectsql("Technical Subjects / Agriculture Stream")
    
ElseIf (SSTab1.Caption = "Aesthetic Subjects") Then
    Call selectsql("Aesthetic Subjects")
    
ElseIf (SSTab1.Caption = "Core Subjects") Then
    Call selectsql("Core Subjects")
    
ElseIf (SSTab1.Caption = "Home Economics Stream") Then
    Call selectsql("Home Economics Stream")
    
ElseIf (SSTab1.Caption = "Religions") Then
    Call selectsql("Religions")
    
End If
End Sub


Public Sub filllist()
On Error Resume Next

Set Rec5 = openDB.OpenRecord("select * from SUBJECT")
    
    
If Not (Rec5.EOF And Rec5.BOF) Then
    Rec5.MoveFirst
        listsubjects.ListItems.clear
        While Not Rec5.EOF
            Set List = listsubjects.ListItems.add
            List.Text = Rec5.Fields(0)
            List.SubItems(1) = Rec5.Fields(1)
            List.SubItems(2) = Rec5.Fields(2)
            Rec5.MoveNext
        Wend
        Else
        listsubjects.ListItems.clear
End If

Rec5.Close
End Sub
