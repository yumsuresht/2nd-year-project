Attribute VB_Name = "basserial"
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
Public serialno As Long

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

Public Function encoding(code As Long) As Long
Dim a As Long
encoding = code + 6111980
End Function
Public Function decoding(decode As Long) As Long
Dim a As Long
decoding = decode - 6111980
End Function

Public Function getserial() As Long
Dim a As Long
Call GetVolumeInfo(Left(App.Path, 3))
serialno = mVolInfo.lVolumeSerialNumber
End Function


Public Sub createfile(st1 As String)
Dim curntLine As String, countLine As Long
Dim s, ln As String
On Error Resume Next
Open App.Path & "\Sch.ini" For Input As #1
Open App.Path & "\temp.txt" For Output As #2
Line Input #1, ln
Print #2, st1
Close 1
Close 2
Kill App.Path & "\Sch.ini"
Name App.Path & "\temp.txt" As App.Path & "\Sch.ini"

openDB.OpenConnection
End Sub
