Attribute VB_Name = "openDB"
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

Global Con As New ADODB.Connection
Public syspaths As String
 

Public Sub OpenConnection()
On Error Resume Next
Dim se, s As String
Dim se1 As Long
syspaths = App.Path

Open App.Path & "\Sch.ini" For Input As #1
If Err.Number = 53 Then
frmserial.Show
Else
Do Until EOF(1)
Line Input #1, se
Loop
Close 1
End If
se1 = se
Call basserial.getserial
If (basserial.decoding(se1) <> basserial.serialno) Then
MsgBox "you must register this software."
frmserial.Show
Exit Sub
End If


Open App.Path & "\School.ini" For Input As #1
If Err.Number = 53 Then
frmODBCLogon.Show
Else
Do Until EOF(1)
Line Input #1, s
Loop
Close 1
End If


ADO:
    Set Con = New Connection
    Con.Open s

If Err.Number = 0 Then
'MsgBox "connection is open !!!", vbInformation

main.Show
frmstafflogin.Show
frmODBCLogon.Hide
Else
MsgBox "Connection Can not open, Set the correct informations !!! " & vbCrLf & _
     Err.Description, vbCritical
frmODBCLogon.Show
End If



'On Error Resume Next
'Set Con = New Connection
'Con.Open "School"
'Con.CursorLocation = adUseClient
'If Err.Number = 0 Then
'MsgBox "connection is open !!!", vbInformation
'main.Show
'frmstafflogin.Show

'Else
'MsgBox "Connection Cannot open, Check the connection !!! " & vbCrLf & _
 '    Err.Description, vbCritical
'End If


End Sub

Public Sub createfile(st As String)
Dim curntLine As String, countLine As Long
Dim s, ln As String
On Error Resume Next

Open App.Path & "\School.ini" For Input As #1
Open App.Path & "\temp.txt" For Output As #2
Line Input #1, ln
Print #2, st
Close 1
Close 2
Kill App.Path & "\School.ini"
Name App.Path & "\temp.txt" As App.Path & "\School.ini"
Call OpenConnection
End Sub


Public Function OpenRecord(RS As String) As Recordset
    Dim R As New ADODB.Recordset
    Set R.ActiveConnection = openDB.Con
    R.CursorLocation = adUseClient
    R.CursorType = adOpenKeyset
    R.LockType = adLockBatchOptimistic
    R.Source = RS
    R.Open
    Set OpenRecord = R
End Function

Public Sub CloseDB()
    Con.Close
End Sub

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
    
End Sub  'GetVo

