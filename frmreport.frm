VERSION 5.00
Begin VB.Form frmreport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4680
   Icon            =   "frmreport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton Option2 
         Caption         =   "Active Student"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Students"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Private Sub Command1_Click()
If (Option1.Value = True) Then
s = "opt1"
Else
s = "opt2"
End If
Call modReports.YearAvgs(s)

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Call modform.FormSize(Me, "Select Report option")
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

