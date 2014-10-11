VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmloding 
   Appearance      =   0  'Flat
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "SCHOOL AUTOMATION SYSTEM"
   ClientHeight    =   4095
   ClientLeft      =   3000
   ClientTop       =   1005
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   3840
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Max             =   500
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8280
      Top             =   3360
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   810
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   7965
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":00C0
      ForeColor       =   &H00FFFF80&
      Height          =   465
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Loding..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   4080
      Left            =   0
      Picture         =   "Form1.frx":0178
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8760
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmloding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

frmloding.Left = (Screen.Width - frmloding.Width) / 2
frmloding.Top = (Screen.Height - frmloding.Height) / 2

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim i As Long
Label2.Caption = Val(Label2.Caption) + 1

For i = 0 To 20
    DoEvents

PBar1.Value = PBar1.Value + i
Next i


If (Label2.Caption = 3) Then

Unload Me
openDB.OpenConnection

End If
End Sub
