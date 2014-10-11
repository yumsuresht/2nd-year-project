VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTimetable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14985
   Icon            =   "frmTimetable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   14985
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   7935
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   14655
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   0
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   51
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   1
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   2
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   49
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   3
         Left            =   9000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   4
         Left            =   11760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   47
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   5
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   6
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   7
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   8
         Left            =   9000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   43
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   9
         Left            =   11760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   10
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   11
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   12
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   13
         Left            =   9000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   14
         Left            =   11760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   15
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   16
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   17
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   18
         Left            =   9000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   19
         Left            =   11760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   20
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   21
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   22
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   23
         Left            =   9000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   24
         Left            =   11760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   25
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   5160
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   26
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   5160
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   27
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   5160
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   28
         Left            =   9000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   5160
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   29
         Left            =   11760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   5160
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   30
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   6000
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   31
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   6000
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   32
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   6000
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   33
         Left            =   9000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   6000
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   34
         Left            =   11760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   6000
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   35
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   6840
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   36
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   6840
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   37
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   6840
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   38
         Left            =   9000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   6840
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   39
         Left            =   11760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   6840
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "MONDAY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   66
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "TUESDAY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   65
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "WEDNESSDAY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   64
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "THURSDAY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9720
         TabIndex        =   63
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "FRIDAY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12720
         TabIndex        =   62
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "INTERVEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   61
         Top             =   3960
         Width           =   9615
      End
      Begin VB.Label Label10 
         Caption         =   "PERIOD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label15 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   6360
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   7080
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
      Begin MSDataListLib.DataCombo dcclass 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Class"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   3615
      End
      Begin MSDataListLib.DataCombo dcStaff 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Teacher ID"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton Option2 
         Caption         =   "Teacher"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Class"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTimetable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecStaff As ADODB.Recordset
Dim RecClass As ADODB.Recordset
Dim RecTime As ADODB.Recordset
Dim RecClassTime As ADODB.Recordset








Private Sub dcclass_Change()
On Error Resume Next
Call displayClasstime(dcclass.Text)
End Sub

Private Sub dcStaff_Change()
Dim staffid As String
staffid = Trim(dcStaff.Text)
    If staffid <> "" Then
    RecStaff.MoveFirst
    RecStaff.Find "StaffID = '" & staffid & "'"
      If RecStaff.EOF Then
            Text1.Text = "Invalid StaffID"
           Call displayStafftime(staffid)

        Else
            Text1.Text = RecStaff!FullName
            Call displayStafftime(staffid)
        End If
        
    End If
End Sub

Private Sub Form_Load()
Call modform.FormSize(Me, "School TimeTable")
Set RecStaff = openDB.OpenRecord("select * from STAFF")
Set RecClass = openDB.OpenRecord("select * from Class")


dcStaff.ListField = "StaffID"
Set dcStaff.RowSource = RecStaff
dcStaff.Text = ""

dcclass.ListField = "ClassName"
Set dcclass.RowSource = RecClass
dcclass.Text = ""
Call Option1_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
RecStaff.Close
RecClass.Close
RecTime.Close
RecClassTime.Close
    main.stbMain.Panels(1).Text = "Status : Not Ready"

End Sub

Private Sub Option1_Click()
Frame3.Visible = True
Frame2.Visible = False
Call ClearTextBoxes(Me)

End Sub

Private Sub Option2_Click()
Frame2.Visible = True
Frame3.Visible = False
Call ClearTextBoxes(Me)

End Sub

Public Function displayStafftime(id As String)
On Error Resume Next
Set RecTime = openDB.OpenRecord("select * from timetable where StaffID='" + id + "' order by Days")

Call ClearTextBoxes(Me)
While Not RecTime.EOF



If (RecTime!Days = "MONDAY") Then
    If (RecTime!Period = 1) Then
    Text2(0).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 2) Then
    Text2(5).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 3) Then
    Text2(10).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 4) Then
    Text2(15).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 5) Then
    Text2(20).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 6) Then
    Text2(25).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 7) Then
    Text2(30).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 8) Then
    Text2(35).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    End If

ElseIf (RecTime!Days = "TUESDAY") Then
    If (RecTime!Period = 1) Then
    Text2(1).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 2) Then
    Text2(6).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 3) Then
    Text2(11).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 4) Then
    Text2(16).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 5) Then
    Text2(21).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 6) Then
    Text2(26).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 7) Then
    Text2(31).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 8) Then
    Text2(36).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    End If
    
ElseIf (RecTime!Days = "WEDNESDAY") Then
    If (RecTime!Period = 1) Then
    Text2(2).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 2) Then
    Text2(7).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 3) Then
    Text2(12).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 4) Then
    Text2(17).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 5) Then
    Text2(22).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 6) Then
    Text2(27).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 7) Then
    Text2(32).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 8) Then
    Text2(37).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    End If

ElseIf (RecTime!Days = "THURSDAY") Then
    If (RecTime!Period = 1) Then
    Text2(3).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 2) Then
    Text2(8).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 3) Then
    Text2(13).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 4) Then
    Text2(18).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 5) Then
    Text2(23).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 6) Then
    Text2(28).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 7) Then
    Text2(33).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 8) Then
    Text2(38).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    End If

ElseIf (RecTime!Days = "FRIDAY") Then
    If (RecTime!Period = 1) Then
    Text2(4).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 2) Then
    Text2(9).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 3) Then
    Text2(14).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 4) Then
    Text2(19).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 5) Then
    Text2(24).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 6) Then
    Text2(29).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 7) Then
    Text2(34).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    ElseIf (RecTime!Period = 8) Then
    Text2(39).Text = RecTime!SubjectNames + "::" + RecTime!ClassName
    End If
End If
RecTime.MoveNext
Wend

RecStaff.Find "StaffID = '" & id & "'"
If RecStaff.EOF Then
   Text1.Text = "Invalid StaffID"
Else
   Text1.Text = RecStaff!FullName
End If
End Function

Public Function displayClasstime(class As String)
On Error Resume Next
Set RecClassTime = openDB.OpenRecord("select T.StaffID,S.FullName,T.Days,T.ClassName,T.SubjectNames,T.Period from timetable T,Staff S where T.StaffID=S.StaffID AND T.ClassName='" + class + "' order by T.Days")

Call ClearTextBoxes(Me)
While Not RecClassTime.EOF



If (RecClassTime!Days = "MONDAY") Then
    If (RecClassTime!Period = 1) Then
    Text2(0).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 2) Then
    Text2(5).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 3) Then
    Text2(10).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 4) Then
    Text2(15).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 5) Then
    Text2(20).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 6) Then
    Text2(25).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 7) Then
    Text2(30).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 8) Then
    Text2(35).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    End If

ElseIf (RecClassTime!Days = "TUESDAY") Then
    If (RecClassTime!Period = 1) Then
    Text2(1).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 2) Then
    Text2(6).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 3) Then
    Text2(11).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 4) Then
    Text2(16).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 5) Then
    Text2(21).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 6) Then
    Text2(26).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 7) Then
    Text2(31).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 8) Then
    Text2(36).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    End If
    
ElseIf (RecClassTime!Days = "WEDNESDAY") Then
    If (RecClassTime!Period = 1) Then
    Text2(2).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 2) Then
    Text2(7).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 3) Then
    Text2(12).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 4) Then
    Text2(17).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 5) Then
    Text2(22).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 6) Then
    Text2(27).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 7) Then
    Text2(32).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 8) Then
    Text2(37).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    End If

ElseIf (RecClassTime!Days = "THURSDAY") Then
    If (RecClassTime!Period = 1) Then
    Text2(3).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 2) Then
    Text2(8).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 3) Then
    Text2(13).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 4) Then
    Text2(18).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 5) Then
    Text2(23).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 6) Then
    Text2(28).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 7) Then
    Text2(33).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 8) Then
    Text2(38).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    End If

ElseIf (RecClassTime!Days = "FRIDAY") Then
    If (RecClassTime!Period = 1) Then
    Text2(4).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 2) Then
    Text2(9).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 3) Then
    Text2(14).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 4) Then
    Text2(19).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 5) Then
    Text2(24).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 6) Then
    Text2(29).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 7) Then
    Text2(34).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    ElseIf (RecClassTime!Period = 8) Then
    Text2(39).Text = RecClassTime!SubjectNames + "::" + RecClassTime!FullName
    End If
End If
RecClassTime.MoveNext
Wend
End Function
