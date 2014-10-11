VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm main 
   BackColor       =   &H80000013&
   Caption         =   "School Automation System"
   ClientHeight    =   5955
   ClientLeft      =   165
   ClientTop       =   540
   ClientWidth     =   8460
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdlHelp 
      Left            =   5160
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5580
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuchangepass 
         Caption         =   "Change password"
      End
      Begin VB.Menu mnustaffAttends 
         Caption         =   "Staff Attendence"
      End
      Begin VB.Menu mnuspa11 
         Caption         =   "-"
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAdministration 
      Caption         =   "&Administration"
      Begin VB.Menu mnuconfigure 
         Caption         =   "Application Setup"
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "User Permissions"
      End
      Begin VB.Menu mnuWizard 
         Caption         =   "Information Wizard"
      End
      Begin VB.Menu mnuspa10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprincipal 
         Caption         =   "Principal"
      End
      Begin VB.Menu mnustaffScheduling 
         Caption         =   "Staff Scheduling"
      End
      Begin VB.Menu mnuRelief 
         Caption         =   "Relief"
      End
      Begin VB.Menu mnuspc11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTimeCreate 
         Caption         =   "Timetable Creater"
      End
   End
   Begin VB.Menu mnuSchool 
      Caption         =   "&Master Setup"
      Begin VB.Menu mnuClass 
         Caption         =   "Class"
      End
      Begin VB.Menu mnuStreams 
         Caption         =   "Streams"
      End
      Begin VB.Menu mnuGradeandSubject 
         Caption         =   "Grade and Subject"
      End
      Begin VB.Menu mnuspa7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClub 
         Caption         =   "Club/Unions"
      End
      Begin VB.Menu mnuspa8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubject 
         Caption         =   "New O/L Subject"
      End
      Begin VB.Menu mnuolsubject 
         Caption         =   "O/L Subject"
      End
      Begin VB.Menu mnuAlsubject 
         Caption         =   "A/L Subject"
      End
   End
   Begin VB.Menu mnustaff 
      Caption         =   "Staff"
      Begin VB.Menu mnunewstaff 
         Caption         =   "Staff Information"
      End
      Begin VB.Menu mnuStaffLeaves 
         Caption         =   "Staff Leaves"
      End
      Begin VB.Menu mnuspa6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclubmaintance 
         Caption         =   "ClubMaintance"
      End
      Begin VB.Menu mnuspa2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTimetable 
         Caption         =   "Time Table"
      End
   End
   Begin VB.Menu mnuStudent 
      Caption         =   "Student"
      Begin VB.Menu mnuNewStudent 
         Caption         =   "New Student"
      End
      Begin VB.Menu mnustuleave 
         Caption         =   "Student Leaving"
      End
      Begin VB.Menu mnutermavg 
         Caption         =   "Term Average"
      End
      Begin VB.Menu mnuMarks 
         Caption         =   "Marks"
      End
      Begin VB.Menu mnuResulte 
         Caption         =   "Results"
         Begin VB.Menu mnuAlresult 
            Caption         =   "A/L Result"
         End
         Begin VB.Menu mnuOlResult 
            Caption         =   "O/L Result"
         End
      End
      Begin VB.Menu mnuspa3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOldstudent 
         Caption         =   "Old Student"
      End
   End
   Begin VB.Menu mnuLibrary 
      Caption         =   "Library"
      Begin VB.Menu mnuBook 
         Caption         =   "New Book"
      End
      Begin VB.Menu mnulibmem 
         Caption         =   "Library Member"
      End
      Begin VB.Menu mnuspace13 
         Caption         =   "-"
      End
      Begin VB.Menu mnulend 
         Caption         =   "Book Lend"
      End
      Begin VB.Menu mnubookreturn 
         Caption         =   "Book Return"
      End
      Begin VB.Menu mnuspac12 
         Caption         =   "-"
      End
      Begin VB.Menu mnupayment 
         Caption         =   "Payment"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Reports"
      Begin VB.Menu mnuLeave 
         Caption         =   "Attendance Report"
      End
      Begin VB.Menu mnuLeaveReport 
         Caption         =   "Staff Leave Report"
      End
      Begin VB.Menu YearAvg 
         Caption         =   "Student Year Average"
      End
      Begin VB.Menu IntAvg 
         Caption         =   "Student Indiual Aveage"
      End
      Begin VB.Menu mnuspac11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOlsheet 
         Caption         =   "O/L Result Sheet"
      End
      Begin VB.Menu mnuALSheet 
         Caption         =   "A/L Result Sheet"
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "Character Certificate"
      End
      Begin VB.Menu mnuspa4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookRept 
         Caption         =   "Book Details"
      End
      Begin VB.Menu mnuOverDue 
         Caption         =   "OverDue Details"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuStatusbar 
         Caption         =   "&Status bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuspa5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTileHorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About SAS"
      End
      Begin VB.Menu mnuCon 
         Caption         =   "Contents..."
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub IntAvg_Click()
'Call modReports.IndYearAvg
frmIndAvg.Show
End Sub

Private Sub MDIForm_Load()
'OpenConnection
    Me.stbMain.Panels(1).Text = "Status: Not ready"


End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

If Cancel = 0 Then
Dim reply
reply = MsgBox("Are you sure you want to exit the program?", vbInformation + vbYesNo, "Exit")
If reply = vbYes Then
Con.Close
End
Else
    Cancel = 1
End If
End If


End Sub



Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuadmin_Click()
frmadmin.Show
End Sub

Private Sub mnuAlresult_Click()
frmAlresult.Show
End Sub

Private Sub mnuALSheet_Click()
frmCharacter.Show
frmCharacter.Command3.Caption = "A/L Result Sheet"

End Sub

Private Sub mnuAlsubject_Click()
frmAlsubject.Show
End Sub

Private Sub mnuBook_Click()
frmaddbook.Show
End Sub

Private Sub mnuBookRept_Click()
Call modReports.Books
End Sub

Private Sub mnubookreturn_Click()
frmreturnbook.Show
End Sub

Private Sub mnuCascade_Click()
Me.Arrange vbCascade

End Sub

Private Sub mnuchangepass_Click()
frmchangepass.Show
End Sub

Private Sub mnuCharacter_Click()
frmCharacter.Show
frmCharacter.Command3.Caption = "Show Character Certificate"
End Sub

Private Sub mnuClass_Click()
frmgrade.Show
End Sub

Private Sub mnuClub_Click()
frmclub.Show
End Sub

Private Sub mnuclubmaintance_Click()
frmclubmaintance.Show
End Sub

Private Sub mnuCon_Click()
 With cdlHelp
 
    .HelpFile = App.Path + "\SASHELP.HLP"
    .HelpCommand = cdlHelpContents
    .ShowHelp
End With
End Sub

Private Sub mnuconfigure_Click()
'frmconfigure.Show
frminitial.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuGradeandSubject_Click()
frmGradeandSubject.Show
End Sub

Private Sub mnuLeave_Click()
Call modReports.StaffAttendance
End Sub

Private Sub mnuLeaveReport_Click()
Call StaffLeavings
End Sub

Private Sub mnulend_Click()
frmlend.Show
End Sub

Private Sub mnulibmem_Click()
frmaddmem.Show
End Sub


Private Sub mnulogoff_Click()
frmstafflogin.Show
End Sub


Private Sub mnuMarks_Click()
frmMarks.Show
End Sub

Private Sub mnunewstaff_Click()
On Error Resume Next
frmnewstaff.Show
End Sub

Private Sub mnuNewStudent_Click()
frmnewstu.Show
End Sub

Private Sub mnuOldstudent_Click()
frmOldboys.Show
End Sub

Private Sub mnuOlResult_Click()
frmOlResult.Show
End Sub

Private Sub mnuOlsheet_Click()
frmCharacter.Show
frmCharacter.Command3.Caption = "O/L Result Sheet"

End Sub

Private Sub mnuolsubject_Click()
frmschsubject.Show
End Sub

Private Sub mnuOverDue_Click()
Call modReports.OverDue

End Sub

Private Sub mnupayment_Click()
frmlibpayment.Show
End Sub

Private Sub mnuprincipal_Click()
frmprincipal.Show
End Sub

Private Sub mnuRelief_Click()
frmRelief.Show
End Sub

Private Sub mnustaffAttends_Click()
frmstaffAttends.Show
End Sub

Private Sub mnuStaffLeaves_Click()
frmstaffleaves.Show
End Sub

Private Sub mnustaffScheduling_Click()
frmstaffScheduling.Show
End Sub

Private Sub mnuStatusbar_Click()
mnuStatusbar.Checked = Not mnuStatusbar.Checked
stbMain.Visible = mnuStatusbar.Checked


End Sub

Private Sub mnuStreams_Click()
frmstreams.Show
End Sub

Private Sub mnustuleave_Click()
frmstulea.Show
End Sub

Private Sub mnuSubject_Click()
frmsubject.Show
End Sub

Private Sub mnutea_subj_Click()
frmtea_subj.Show
End Sub

Private Sub mnutermavg_Click()
frmtermavg.Show
End Sub

Private Sub mnuTileHorizontally_Click()
Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVertically_Click()
Me.Arrange vbTileVertical

End Sub

Private Sub mnuTimeCreate_Click()
frmTimetableCreater.Show
End Sub

Private Sub mnuTimetable_Click()
frmTimetable.Show
End Sub

Private Sub mnuWizard_Click()
frmWizard.Show
End Sub

Private Sub YearAvg_Click()
frmreport.Show
End Sub
