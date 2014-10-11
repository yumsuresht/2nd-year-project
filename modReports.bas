Attribute VB_Name = "modReports"
Public RecStaffLea As ADODB.Recordset
Public RecStaffLea1 As ADODB.Recordset
Public RecOLSheet As ADODB.Recordset
Public RecBook As ADODB.Recordset
Public RecBook1 As ADODB.Recordset
Public RecYearAvg As ADODB.Recordset
Public RecIndYearAvg As ADODB.Recordset
Public RecSchool As ADODB.Recordset







Public Sub StaffAttendance()
On Error Resume Next
    Set RecStaffLea = openDB.OpenRecord("select * from STAFFATTENDANCE  order by attDate,ComingTime asc")
    Set RecSchool = openDB.OpenRecord("select * from SCHOOL")

    With drStaffLeaves
        Set .DataSource = RecStaffLea
        .Caption = "Staff Attendance Details"
        .LeftMargin = 750
        .Sections("Section4").Controls("Label17").Caption = RecSchool!SCHOOLNAME
        .Sections("Section4").Controls("Label2").Caption = RecSchool!ADDRESS1 + " ," + RecSchool!ADDRESS2


        .Sections("Section1").Controls("Text1").DataField = "StaffID"
        .Sections("Section1").Controls("Text2").DataField = "AttDate"
        .Sections("Section1").Controls("Text3").DataField = "ComingTime"
        .Sections("Section1").Controls("Text4").DataField = "GoingTime"
        .WindowState = vbMaximized
        .Show
    End With
End Sub

Public Sub StaffLeavings()
On Error Resume Next

    Set RecStaffLea1 = openDB.OpenRecord("select S.StaffID,A.FullName,S.DateTo,S.DateFrom,S.LeaveType from STAFFLEAVES S,STAFF A where S.StaffID=A.StaffID")
    Set RecSchool = openDB.OpenRecord("select * from SCHOOL")

    With drStaffLeaves1
        Set .DataSource = RecStaffLea1
        .Caption = "Staff Leaving Details"
        .LeftMargin = 750
        .Sections("Section4").Controls("Label17").Caption = RecSchool!SCHOOLNAME
        .Sections("Section4").Controls("Label7").Caption = RecSchool!ADDRESS1 + " ," + RecSchool!ADDRESS2

        
        .Sections("Section1").Controls("Text1").DataField = "StaffID"
        .Sections("Section1").Controls("Text2").DataField = "FullName"
        .Sections("Section1").Controls("Text3").DataField = "DateFrom"
        .Sections("Section1").Controls("Text4").DataField = "DateTo"
        .Sections("Section1").Controls("Text5").DataField = "LeaveType"
        .WindowState = vbMaximized
        .Show
    End With
End Sub

Public Sub OlResultSheet(qu As String)
On Error Resume Next
    Set RecOLSheet = openDB.OpenRecord("select * from OLRESULT O,MAINSTUDENTS M where O.StuID=M.StuID and O.StuID='" + qu + "'")
        Set RecSchool = openDB.OpenRecord("select * from SCHOOL")

    With drOLresult
        Set .DataSource = RecOLSheet
        .Caption = "O/L Result Sheet"
        .LeftMargin = 750
        
        .Sections("Section4").Controls("Label17").Caption = RecSchool!SCHOOLNAME
        .Sections("Section4").Controls("Label27").Caption = RecSchool!ADDRESS1 + " ," + RecSchool!ADDRESS2

        .Sections("Section2").Controls("Label20").Caption = RecOLSheet.Fields(15) + " " + RecOLSheet.Fields(14)

        .Sections("Section1").Controls("Text1").DataField = "OlYear"
        .Sections("Section1").Controls("Text2").DataField = "IndexNo"
               
        .Sections("Section1").Controls("Text11").DataField = "Maths"
        .Sections("Section1").Controls("Text12").DataField = "Science"
        .Sections("Section1").Controls("Text13").DataField = "Social"
        .Sections("Section1").Controls("Text14").DataField = "English"
        
        .Sections("Section1").Controls("Label6").Caption = UCase(Left(RecOLSheet.Fields(7), Len(RecOLSheet.Fields(7)) - 2))
        .Sections("Section1").Controls("Label5").Caption = Right(RecOLSheet.Fields(7), 1)
        .Sections("Section1").Controls("Label7").Caption = UCase(Left(RecOLSheet.Fields(8), Len(RecOLSheet.Fields(8)) - 2))
        .Sections("Section1").Controls("Label9").Caption = Right(RecOLSheet.Fields(8), 1)
        .Sections("Section1").Controls("Label8").Caption = UCase(Left(RecOLSheet.Fields(9), Len(RecOLSheet.Fields(9)) - 2))
        .Sections("Section1").Controls("Label10").Caption = Right(RecOLSheet.Fields(9), 1)

        .Sections("Section1").Controls("Label11").Caption = UCase(Left(RecOLSheet.Fields(10), Len(RecOLSheet.Fields(10)) - 2))
        .Sections("Section1").Controls("Label12").Caption = Right(RecOLSheet.Fields(10), 1)
        .Sections("Section1").Controls("Label13").Caption = UCase(Left(RecOLSheet.Fields(11), Len(RecOLSheet.Fields(11)) - 2))
        .Sections("Section1").Controls("Label14").Caption = Right(RecOLSheet.Fields(11), 1)
        .Sections("Section1").Controls("Label15").Caption = UCase(Left(RecOLSheet.Fields(12), Len(RecOLSheet.Fields(12)) - 2))
        .Sections("Section1").Controls("Label16").Caption = Right(RecOLSheet.Fields(12), 1)

        .WindowState = vbMaximized
        .Show
    End With
End Sub


Public Sub Books()
On Error Resume Next

    Set RecBook = openDB.OpenRecord("select * from BOOK")
    Set RecSchool = openDB.OpenRecord("select * from SCHOOL")

    With drBook
        Set .DataSource = RecBook
        .Caption = "Book Details"
        .LeftMargin = 750
                
        .Sections("Section4").Controls("Label6").Caption = RecSchool!SCHOOLNAME
        .Sections("Section4").Controls("Label8").Caption = RecSchool!ADDRESS1 + " ," + RecSchool!ADDRESS2


        
        .Sections("Section1").Controls("Text1").DataField = "BookID"
        .Sections("Section1").Controls("Text3").DataField = "Title"
        .Sections("Section1").Controls("Text4").DataField = "AuthorName"
        .Sections("Section1").Controls("Text5").DataField = "Catagory"
        .Sections("Section1").Controls("Text6").DataField = "N_OF_Co"
        .WindowState = vbMaximized
        .Show
    End With
End Sub


Public Sub OverDue()
On Error Resume Next

    Set RecBook1 = openDB.OpenRecord("select distinct(LE.MemID),L.MemName,L.Status,B.Title,LE.BorrowDate,LE.DueDate from LENDING LE,LIBRARYMEMBER L,COPY_OF_BOOK C ,BOOK B where C.BookID=B.BookID and C.AccessNo=LE.AccessNo and L.MemID=LE.MemID and LE.Duedate<getDate()")
    Set RecSchool = openDB.OpenRecord("select * from SCHOOL")

    With drBookDue
        Set .DataSource = RecBook1
        .Caption = "Over Due Details"
        .LeftMargin = 750
                
        .Sections("Section4").Controls("Label9").Caption = RecSchool!SCHOOLNAME
        .Sections("Section4").Controls("Label12").Caption = RecSchool!ADDRESS1 + " ," + RecSchool!ADDRESS2

        .Sections("Section1").Controls("Text1").DataField = "MemID"
        .Sections("Section1").Controls("Text2").DataField = "MemName"
        .Sections("Section1").Controls("Text3").DataField = "Status"
        .Sections("Section1").Controls("Text4").DataField = "Title"
        .Sections("Section1").Controls("Text5").DataField = "BorrowDate"
        .Sections("Section1").Controls("Text6").DataField = "DueDate"
        .Sections("Section3").Controls("Label7").Caption = Now
        .Sections("Section3").Controls("Label10").Caption = modform.StaffName

        .WindowState = vbMaximized
        .Show
    End With
End Sub

Public Sub YearAvgs(opt As String)
On Error Resume Next
    Set RecSchool = openDB.OpenRecord("select * from SCHOOL")

If (opt = "opt1") Then
    Set RecYearAvg = openDB.OpenRecord("select M.StuID,M.StudentName,M.FatherName,Year6,Year7,Year8,Year9,Year10,Year11,Year12,Year13 from MAINSTUDENTS M,YEARAVERAGE Y where Y.StuID=M.StuID")
ElseIf (opt = "opt2") Then
    Set RecYearAvg = openDB.OpenRecord("select M.StuID,M.StudentName,M.FatherName,Year6,Year7,Year8,Year9,Year10,Year11,Year12,Year13 from MAINSTUDENTS M,YEARAVERAGE Y, ACTIVESTUDENT A where Y.StuID=M.StuID and A.StuID=M.StuID")
End If


    'Set RecYearAvg = openDB.OpenRecord("select M.StuID,M.StudentName,Year6,Year7,Year8,Year9,Year10,Year11,Year12,Year13 from MAINSTUDENTS M,YEARAVERAGE Y where Y.StuID=M.StuID")
    With drYearAvg
        Set .DataSource = RecYearAvg
        .Caption = "Student Year Average"
        .LeftMargin = 750
        
        .Sections("Section4").Controls("Label17").Caption = RecSchool!SCHOOLNAME
        .Sections("Section4").Controls("Label24").Caption = RecSchool!ADDRESS1 + " ," + RecSchool!ADDRESS2

        .Sections("Section1").Controls("Text1").DataField = "StuID"
        .Sections("Section1").Controls("Text2").DataField = "StudentName"
        .Sections("Section1").Controls("Text11").DataField = "FatherName"
        
        .Sections("Section1").Controls("Text3").DataField = "Year6"
        .Sections("Section1").Controls("Text4").DataField = "Year7"
        .Sections("Section1").Controls("Text5").DataField = "Year8"
        .Sections("Section1").Controls("Text6").DataField = "Year9"
        .Sections("Section1").Controls("Text7").DataField = "Year10"
        .Sections("Section1").Controls("Text8").DataField = "Year11"
        .Sections("Section1").Controls("Text9").DataField = "Year12"
        .Sections("Section1").Controls("Text10").DataField = "Year13"
        .Sections("Section3").Controls("Label14").Caption = Now
        .Sections("Section3").Controls("Label12").Caption = modform.StaffName
        .Sections("Section3").Controls("Label22").Caption = RecYearAvg.RecordCount
        
        .WindowState = vbMaximized
        .Show
    End With
    
End Sub

Public Sub IndYearAvg(s As String)
On Error Resume Next

    Set RecIndYearAvg = openDB.OpenRecord("select M.StuID,M.StudentName,M.FatherName,M.D_Of_Admin,M.AdminGrade,Year6,Year7,Year8,Year9,Year10,Year11,Year12,Year13 from MAINSTUDENTS M,YEARAVERAGE Y where Y.StuID=M.StuID and Y.StuID='" & s & "'")
        Set RecSchool = openDB.OpenRecord("select * from SCHOOL")

    With drIndualPer
        Set .DataSource = RecIndYearAvg
        .Caption = "Indiual Performance Report"
        .LeftMargin = 750
        .Sections("Section4").Controls("Label9").Caption = RecSchool!SCHOOLNAME
        .Sections("Section4").Controls("Label27").Caption = RecSchool!ADDRESS1 + " ," + RecSchool!ADDRESS2

        .Sections("Section1").Controls("Text1").DataField = "StuID"
        .Sections("Section1").Controls("Text2").DataField = "StudentName"
        .Sections("Section1").Controls("Text3").DataField = "FatherName"
        .Sections("Section1").Controls("Text4").DataField = "D_Of_Admin"
        .Sections("Section1").Controls("Text5").DataField = "AdminGrade"
        .Sections("Section1").Controls("Text6").DataField = "Year6"
        .Sections("Section1").Controls("Text7").DataField = "Year7"
        .Sections("Section1").Controls("Text8").DataField = "Year8"
        .Sections("Section1").Controls("Text9").DataField = "Year9"
        .Sections("Section1").Controls("Text10").DataField = "Year10"
        .Sections("Section1").Controls("Text11").DataField = "Year11"
        .Sections("Section1").Controls("Text12").DataField = "Year12"
        .Sections("Section1").Controls("Text13").DataField = "Year13"
        .Sections("Section1").Controls("Label22").Caption = modform.StaffName
        .Sections("Section1").Controls("Label25").Caption = Now


        .WindowState = vbMaximized
        .Show
    End With
End Sub
