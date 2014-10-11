Attribute VB_Name = "modform"
Global userid As String
Global Staff1 As String
Global StaffName As String

Global formname As String
Dim rec1 As ADODB.Recordset



Public Sub FormSize(Frm As Form, title As String)
    Frm.Left = (main.ScaleWidth - Frm.Width) / 2
    Frm.Top = (main.ScaleHeight - Frm.Height) / 2
    Frm.Caption = UCase(title)
    main.stbMain.Panels(1).Text = "Status : Ready"

End Sub

Public Sub ClearTextBoxes(Frm As Form)
    Dim Ctrl As Control
    For Each Ctrl In Frm.Controls
        If TypeOf Ctrl Is TextBox Then
            Ctrl.Text = ""
        End If
    Next Ctrl
End Sub

Public Sub uppercase(Frm As Form)
    Dim Ctrl As Control
    For Each Ctrl In Frm.Controls
        If TypeOf Ctrl Is TextBox Then
            Ctrl.Text = UCase(Ctrl.Text)
        End If
    Next Ctrl
End Sub

Public Function setUserID(Uid As String)
On Error Resume Next
Set rec1 = openDB.OpenRecord("select FullName from STAFF where StaffID='" & Uid & "'")
If (rec1.RecordCount = 0) Then
StaffName = Uid
Else
StaffName = rec1!FullName
End If
userid = Uid
End Function
