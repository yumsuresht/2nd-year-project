VERSION 5.00
Begin VB.Form frmconfigure 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ID Configure"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   Icon            =   "frmconfigure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6480
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   375
         Left            =   4440
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Warning..."
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
         Left            =   2400
         TabIndex        =   18
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "You can enter value only one time by this window, So please fill carefully"
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
         Height          =   495
         Left            =   2400
         TabIndex        =   17
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Student ID"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Staff ID"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Library MemberID"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Library Receipt No"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Book ReserveNo"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Library Transaction No"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Temp ID"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmconfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset

Private Sub Command1_Click()
On Error Resume Next
        rec.AddNew
        rec!TEMIDS = Text1.Text
        rec!STUIDS = Text2.Text
        rec!STAFFIDS = Text3.Text
        rec!MEMBERIDS = Val(Text4.Text)
        rec!RECEPID = Val(Text5.Text)
        rec!RESERVENO = Val(Text6.Text)
        rec!LIBTRANSNO = Val(Text7.Text)
        rec.UpdateBatch
    MsgBox ("Sucessfully added")
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub




Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "Application Setup")
    Set rec = openDB.OpenRecord("SELECT * FROM IDS")
    
    If (rec!TEMIDS = "") Then
    Command1.Enabled = True
    Else
    Command1.Enabled = False
    End If
    main.Enabled = False
    
    
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
main.Enabled = True
rec.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

