VERSION 5.00
Begin VB.Form frmserial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Registration"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Submit Information"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Text            =   "SAS"
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Enter your Serial Number received from SAS. Please make sure it is all letters are correct"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmserial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Listing 25.41
Private Type VolumeInfoType
  sVolName As String * 255
  lBufferSize As Long
  lVolumeSerialNumber As Long
  lMaxFileLength As Long
  lFileSysFlags As Long
  sFileSysName As String * 255
  lFileSysBufSize As Long
End Type

Dim mVolInfo As VolumeInfoType
Private Sub Command1_Click()
On Error Resume Next
  Call GetVolumeInfo(Left(App.Path, 3))

If (Text1.Text <> "SAS") Then
MsgBox "Incorrect Serial Number. Please input the correct lenght including the hyphens", vbExclamation
Exit Sub

ElseIf (decoding(Text2.Text) <> mVolInfo.lVolumeSerialNumber) Then
MsgBox "Incorrect Serial Number. Please input the correct lenght including the hyphens", vbExclamation
Exit Sub

Else
Call basserial.createfile(Text2.Text)
Unload Me
End If
End Sub

Public Function encoding(code As Long) As Long
Dim a As Long
encoding = code + 6111980
End Function
Public Function decoding(decode As Long) As Long
Dim a As Long
decoding = decode - 6111980
End Function

Sub GetVolumeInfo(sPathName As String)
  Dim lRetValue As Long
  
  mVolInfo.sVolName = String(255, " ")
  mVolInfo.sFileSysName = String(255, " ")
  
  mVolInfo.lBufferSize = Len(mVolInfo.sVolName)
  mVolInfo.lFileSysBufSize = Len(mVolInfo.sFileSysName)
  
  lRetValue = apiGetVolumeInformation( _
     sPathName, _
     mVolInfo.sVolName, _
     mVolInfo.lBufferSize, _
     mVolInfo.lVolumeSerialNumber, _
     mVolInfo.lMaxFileLength, _
     mVolInfo.lFileSysFlags, _
     mVolInfo.sFileSysName, _
     mVolInfo.lFileSysBufSize)
    
End Sub  'GetVolumeInfo

Private Sub Command3_Click()
On Error Resume Next
  Call GetVolumeInfo(Left(App.Path, 3))
'Text3.Text = mVolInfo.lVolumeSerialNumber

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call modform.FormSize(Me, "System Register")

End Sub

