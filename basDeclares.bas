Attribute VB_Name = "basDeclares"
Option Explicit

'== Begin Public Constant Declarations ===================

'Constants used throughout
Public Const MAX_PATH = 260
Public Const MAX_COMPUTERNAME_LENGTH = 15


'Listing 25.27
'Constant declarations to use with the GetVersion API call.
Public Const VER_PLATFORM_WIN32S = 0
Public Const VER_PLATFORM_WIN32_Windows = 1
Public Const VER_PLATFORM_WIN32_NT = 2


'Listing 25.34
'Constant values to use with GetDriveType.
Public Const DRIVE_UNKNOWN = 0
Public Const DRIVE_NOT_AVAILABLE = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6


'== Begin Public Type Declarations =======================

'Listing 25.25
'OSVERSIONINFO type. This structure is used by
'the GetVersionEx to retrieve information about
'the Windows version installed on the local computer.
Public Type OSVERSIONINFO
  lVersionInfo As Long
  lMajorVersion As Long
  lMinorVersion As Long
  lBuildNumber As Long
  lplatformID As Long
  sVersion As String * 128
End Type 'OSVERSIONINFO


'== Begin Public API Function Declarations ===============


'The following declare is not discussed
'in Chapter 25, but shows how to control
'the system speaker with API calls.
Declare Function apiBeep Lib "Kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long


'Listing 25.7
'The GetWindowText API function prototype.
Declare Function apiGetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal sCaption As String, ByVal lCaptionSize As Long) As Long

'Listing 25.8
'The GetWindowTextLengthA API function prototype.
Declare Function apiGetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long


'Listing 25.10
'The SetWindowText API function prototype.
Declare Function apiSetWindowText Lib "User32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal sCaption As String) As Long


'Listing 25.13
'The GetParent API function prototype.
Declare Function apiGetParent Lib "User32" Alias "GetParent" (ByVal hWnd As Long) As Long


'Listing 25.15
'The GetCommandLine API function prototype.
Declare Function apiGetCommandLine Lib "Kernel32" Alias "GetCommandLineA" () As String


'Listing 25.17
'The GetClassName API function prototype.
Declare Function apiGetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal sClassName As String, ByVal lClassSize As Long) As Long


'Listing 25.19
'The GetWindowsDirectoryA API function prototype.
Declare Function apiGetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal sReturnBuffer As String, ByVal lBuffSize As Long) As Long


'Listing 25.23
'Using the ByVal option.
Declare Function apiGetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal lBufferSize As Long, ByVal sReturnBuffer As String) As Long


'Listing 25.21
'The GetSystemDirectory API function prototype.
'sReturnBuffer is the system directory path
'lBufferSize is the size of sReturnBuffer when passed in
'Returns the number of bytes returned in sReturnBuffer
Declare Function apiGetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal sReturnBuffer As String, ByVal lBufferSize As Long) As Long


'Listing 25.23
'The GetDiskFreeSpace API function prototype.
Declare Function apiGetDiskFreeSpace Lib "Kernel32" Alias "GetDiskFreeSpaceA" (ByVal sPath As String, lSectors As Long, lBytes As Long, lFreeClusters As Long, lClusters As Long) As Long


'Listing 25.25
'The GetVolumeInformation API function prototype.
Declare Function apiGetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal sPath As String, ByVal sNameBuffer As String, ByVal lVolumeNameSize As Long, lVolSerialNo As Long, lMaxFileLength As Long, lSystemFlags As Long, ByVal sSysNamebuffer As String, ByVal lSysNameBufSize As Long) As Long


'Listing 25.26
'The GetVersion function prototype.
Declare Function apiGetVersion Lib "Kernel32" Alias "GetVersionExA" (ByRef osVer As OSVERSIONINFO) As Long


'Listing 25.29
'The GetUserName API function prototype.
Declare Function apiGetUserName Lib "Advapi32" Alias "GetUserNameA" (ByVal sBuffer As String, lBufferSize As Long) As Long


'Listing 25.31
'The GetComputerName API function prototype.
Declare Function apiGetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lBufferSize As Long) As Long


'Listing 25.33
'The GetDriveType API function prototype.
Declare Function apiGetDriveType Lib "Kernel32" Alias "GetDriveTypeA" (ByVal sPath As String) As Long



