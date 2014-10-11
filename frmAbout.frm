VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SAS"
   ClientHeight    =   3840
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5955
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2650.436
   ScaleMode       =   0  'User
   ScaleWidth      =   5592.054
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   3960
      Top             =   1320
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "Show Developers"
         Height          =   450
         Left            =   4320
         TabIndex        =   7
         Top             =   3000
         Width           =   1260
      End
      Begin VB.CommandButton cmdSysInfo 
         Caption         =   "&System Info..."
         Height          =   345
         Left            =   4320
         TabIndex        =   2
         Top             =   2520
         Width           =   1260
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   345
         Left            =   4320
         TabIndex        =   1
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmAbout.frx":030A
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblDisclaimer 
         Caption         =   $"frmAbout.frx":03F1
         Height          =   825
         Left            =   270
         TabIndex        =   6
         Top             =   2160
         Width           =   3630
      End
      Begin VB.Label lblTitle 
         Caption         =   "SCHOOL AUTOMATION SYSTEM (SAS)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   5445
      End
      Begin VB.Label lblDescription 
         Caption         =   $"frmAbout.frx":04BE
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   150
         TabIndex        =   4
         Top             =   600
         Width           =   5355
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   0
         X2              =   5564
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblVersion 
         Height          =   255
         Left            =   270
         TabIndex        =   3
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5564
         Y1              =   1800
         Y2              =   1800
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   2040
      Left            =   240
      TabIndex        =   8
      Top             =   5506
      Width           =   5535
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Group Member"
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   3360
         TabIndex        =   17
         Top             =   1250
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   $"frmAbout.frx":0598
         ForeColor       =   &H0000FFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "Group Leader"
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "Group Member"
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   3360
         TabIndex        =   14
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Group Member"
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   3360
         TabIndex        =   13
         Top             =   650
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "Group Member"
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   3360
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Group Member"
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   3360
         TabIndex        =   11
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Group Member"
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   3360
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "e-mail: groupc38@yahoo.com"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   2415
      End
   End
   Begin VB.Line Line2 
      X1              =   225.372
      X2              =   5296.253
      Y1              =   2567.61
      Y2              =   2567.61
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
Dim i, j As Long

If (Command1.Caption = "Show Developers") Then
frmAbout.Height = 6100
Command1.Caption = "Hide Developers"
Timer1.Enabled = True

ElseIf (Command1.Caption = "Hide Developers") Then
frmAbout.Height = 4300
Command1.Caption = "Show Developers"
Timer1.Enabled = False

End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call modform.FormSize(Me, "About")

    Me.Caption = "About " & App.title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.title
    main.Enabled = False
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Unload(Cancel As Integer)
    main.Enabled = True
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Frame1_Click()
Unload Me
End Sub

Private Sub Frame2_Click()
Timer1.Enabled = False
End Sub

Private Sub Frame2_DblClick()
Timer1.Enabled = True

End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label16_Click()
Timer1.Enabled = False

End Sub

Private Sub Label16_DblClick()
Timer1.Enabled = True

End Sub

Private Sub Label9_Click()
Timer1.Enabled = False

End Sub

Private Sub Label9_DblClick()
Timer1.Enabled = True

End Sub

Private Sub lblDescription_Click()
Unload Me
End Sub

Private Sub lblDisclaimer_Click()
Unload Me
End Sub

Private Sub lblTitle_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Dim i, j As Long
Frame2.Top = Frame2.Top - 1
If (Frame2.Top < 2500 And Frame2.Top > 2450) Then
Frame2.Top = 3800
End If
End Sub
