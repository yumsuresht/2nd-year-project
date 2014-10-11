select Distinct(B.Catagory)from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID
SELECT A.StuID,M.StudentName,M.FatherName,A.Curr_Class FROM MAINSTUDENTS M,ACTIVESTUDENT A where M.TemID=A.TemID and Old_Status='No'
select StaffID,FullName,Work_Hours from STAFF where PostHeld='TEACHER'
select * from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID and C.Status='LEND'
select B.BookID,C.AccessNo,B.Title,B.Edition,B.Catagory AS Category,C.Status from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID and accessno='455-1'
select * from RESERVE R,COPY_OF_BOOK C,BOOK B where C.AccessNo=R.AccessNo and B.BookID=C.BookID
select Curr_Class from LIBRARYMEMBER L,ACTIVESTUDENT A where L.SCID=A.StuID
select * from BOOK where Title  Like '%gh%'
select * from ACTIVESTUDENT WHERE CURR_CLASS LIKE '11 %'
select * from CLASS where ClassName like '12 %'
select * from STREAM
select ReserveDate from RESERVE where AccessNo='3000-2'
select * from LIBRARYMEMBER where LendStatus IN('YES','RES')
select * from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID and C.Status='LEND'
select L.AccessNo,B.Title,L.MemID,M.MemName,L.BorrowDate,L.DueDate 
from LENDING L,COPY_OF_BOOK C,Book B ,LIBRARYMEMBER M
where L.AccessNo=C.AccessNo 
and B.BookID=C.BookID 
and M.MemID= L.MemID
and L.TRANSACT='LEND'
select *
from LENDING L,COPY_OF_BOOK C,Book B ,LIBRARYMEMBER M
where L.AccessNo=C.AccessNo 
and B.BookID=C.BookID 
and M.MemID= L.MemID
and C.BookStatus='LEND' and M.LendStatus='LEND'
select * from LENDING
select * from COPY_OF_BOOK
select * from LIBRARYMEMBER M where M.LendStatus='LEND'
select B.BookID,C.AccessNo,B.Title,B.ISBN,B.Catagory AS Category from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID and C.AccessNo='1000-1'
select * from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID and C.AccessNo='1000-3'
select * from COPY_OF_BOOK where AccessNo='1000-1'
select * from SUBJECT where Category= 'Additional Subjects'
select * from OLSUBJECT where Category= 'Additional Subjects' and Status='YES'
select * from OLSUBJECT where Category= 'Additional Subjects' and Status='NO'
select * from OLSUBJECT where Category= 'Additional Subjects' and SubjectNames='Tamil Lit'
SELECT * FROM MAINSTUDENTS M,ACTIVESTUDENT A where M.TemID=A.TemID and Old_Status='No'
select DISTINCT(Category) from SUBJECT
select * from SUBJECT where Category= 'Additional Subjects' and SubjectNames='Sanskrit'
select * from CLUB
select * from CLUBMEMBER where CName='TAMIL UNION'
SELECT A.StuID,M.StudentName,M.FatherName,A.Curr_Class FROM MAINSTUDENTS M,ACTIVESTUDENT A where M.TemID=A.TemID and Old_Status='No'
SELECT A.StuID,M.StudentName,M.FatherName,A.Curr_Class FROM MAINSTUDENTS M,ACTIVESTUDENT A,CLUBMEMBER C where M.TemID=A.TemID and Old_Status='No' and C.CName='TAMIL UNION' and A.StuID=C.StuID
SELECT * FROM MAINSTUDENTS M,ACTIVESTUDENT A,CLUBMEMBER C where M.TemID=A.TemID and Old_Status='No'and A.StuID=C.StuID and C.CName='TAMIL UNION' 
SELECT A.StuID,M.StudentName,M.FatherName FROM MAINSTUDENTS M,ACTIVESTUDENT A,CLUBMAINTAINCE C,STAFF S where M.TemID=A.TemID and Old_Status='No'and A.StuID=C.Pres_StuID and C.CName='TAMIL UNION' 
SELECT A.StuID,M.StudentName,M.FatherName FROM MAINSTUDENTS M,ACTIVESTUDENT A,CLUBMAINTAINCE C,STAFF S where M.TemID=A.TemID and Old_Status='No'and A.StuID=C.Sec_StuID and C.CName='TAMIL UNION' 
select S.StaffID,S.FullName from CLUBMAINTAINCE C,STAFF S where S.StaffID=C.Super_StaffID and C.CName='TAMIL UNION' 
select * from CLUBMAINTAINCE where CName='TAMIL UNION' 
select * from TERMAVG
select * from YEARAVERAGE
SELECT A.StuID,M.StudentName,A.Curr_Class,T.Term1,T.Term2,T.Term3 FROM MAINSTUDENTS M,ACTIVESTUDENT A,TERMAVG T where M.TemID=A.TemID and Old_Status='No'and T.StuID=A.StuID
SELECT A.StuID,M.StudentName,M.FatherName,A.Curr_Class FROM MAINSTUDENTS M,ACTIVESTUDENT A where M.TemID=A.TemID and Old_Status='No' and Curr_Class='8 A'
select T.StuID,T.Grade,T.Term1,T.Term2,T.Term3 from TERMAVG T where T.Grade='8 A'
select * from CLASS
select * from TERMAVG
SELECT A.StuID,M.StudentName,M.FatherName FROM MAINSTUDENTS M,ACTIVESTUDENT A,TERMAVG T where M.TemID=A.TemID and Old_Status='No' and T.StuID=A.StuID and T.Grade='8 A'
select * from ACTIVESTUDENT where Old_Status='No'
select Curr_Class,COUNT(StuID)AS Total from ACTIVESTUDENT where Old_Status='No' group by Curr_Class
select COUNT(StuID)AS Total from ACTIVESTUDENT where Old_Status='No'AND Curr_Class LIKE '7%'
select COUNT(StuID)AS Total from ACTIVESTUDENT where Old_Status='No' and Curr_Class LIKE '6 A'
select distinct(Category) from OLSUBJECT
select count(*) from CLASS where ClassName LIKE '10 %'
select count(*) from CLASS where ClassName LIKE '11 %'
select COUNT(StuID)AS Total from ACTIVESTUDENT where Old_Status='No' and Curr_Class LIKE '10 A'
select * from ACTIVESTUDENT where Curr_Class like '7 %'
select A.StuID,M.StudentName,M.FatherName,A.Curr_Class  from MAINSTUDENTS M,ACTIVESTUDENT A where M.StuID=A.StuID
select *  from MAINSTUDENTS M,ACTIVESTUDENT A where M.StuID=A.StuID
select * from MAINSTUDENTS M,ACTIVESTUDENT A,CLUBMEMBER c WHERE M.StuID=A.StuID AND A.StuID=C.StuID
SELECT * FROM CLUBMAINTAINCE
select * from LIBRARYMEMBER L WHERE L.LENDSTATUS IN('LEND','FINE')
select * from ACTIVESTUDENT WHERE CURR_CLASS IN('11 R1','11 R2')
select * from TEMPSTUDENTS where TemID='1020'
select * from TEMPSTUDENTS where StudentName like 'subas'
select * from TEMPSTUDENTS where FatherName like 'karan'
select * from TEMPSTUDENTS where Street like '%jaffna%' or City like '%jaffna%'
select * from ALRESULT
select * from FOLLOWSTREAM
select  A.SUBJECT1 ,B.SUBJECT1,A.SUBJECT2 ,B.SUBJECT2,A.SUBJECT3 ,B.SUBJECT3 from  FOLLOWSTREAM A, Alresult B WHERE  A.SID=b.SID AND A.SID=1000
select * from ALSUBJECT WHERE STREAM='COMMERCE'
select * from FOLLOWSTREAM F, ACTIVESTUDENT A WHERE F.StuID=A.StuID
select * from ALRESULT A,FOLLOWSTREAM F WHERE A.STUID=F.STUID AND F.STUID='5244' order by Alyear
select * from ALRESULT where StuId='5244'
select * from LOGIN where StaffID='5011'
select * from STAFFATTENDANCE where goingtime is null and StaffID='5010'
select * from SHORTLEAVES where StaffID='5010' and month(dates)=month(GETDATE()) and Descriptions='Short Leave'
select * from SHORTLEAVES
select * from SHORTLEAVES where ComingTime is null
select * from SHORTLEAVES where ComingTime is null and Dates= '2004-09-21' & "' and StaffID='" & userid & "'")
select distinct(SubjectNames) from SUBJECT union select distinct(SubjectNames) from ALSUBJECT 
select * from MAINSTUDENTS
select * from ACTIVESTUDENT
select A.StuID,M.StudentName from ACTIVESTUDENT A,MAINSTUDENTS M WHERE A.StuID=M.StuID and A.Curr_Class LIKE '11%'
select * from ALSUBJECT
select StaffID,FullName from STAFF
select * from OLSUBJECT where Category='Additional Subjects'
select * from OLSUBJECT where Category='Aesthetic Subjects'
select * from OLSUBJECT where Category='Core Subjects' and SubjectNames LIKE '%Language%'
select * from OLSUBJECT where Category='Religions'
select SubjectNames from OLSUBJECT where Category IN ('Commerce Stream','Home Economics Stream','Technical Stream','Technical Subjects / Agriculture Stream')
select * from OLRESULT
select CName,JoinDate,StuID,'Mem' AS POST from CLUBMEMBER where STUID='5246'  order by CName asc,Post desc
select CName,Years,Pres_StuID,'Pres' AS POST from CLUBMAINTAINCE 
select CName,Years,Sec_StuID,'Sec' AS POST from CLUBMAINTAINCE where Sec_StuID='5246'
select * from MAINSTUDENTS
select * from OLDBOYS
select M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin AS Admission_Date,O.D_Of_Leave AS Leave_Date
from MAINSTUDENTS M, OLDBOYS O where M.StuID=O.StuID and M.StuID='5248'
select M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin AS Admission_Date,O.D_Of_Leave AS Leave_Date
from MAINSTUDENTS M, OLDBOYS O where M.StuID=O.StuID and M.Street='vbvb'
select * from OLRESULT WHERE STUID='5250'
select * from ALRESULT
select * from STAFFATTENDANCE
select * from STAFF
select * from STAFFATTENDANCE
select * from SHORTLEAVES
select * from MEDICALLEAVES
select * from STAFFLEAVES
select * from CASUALLEAVES
select * from STAFF
select S.StaffID,A.FullName,S.DateTo,S.DateFrom,S.LeaveType from STAFFLEAVES S,STAFF A where S.StaffID=A.StaffID

select distinct(M.StuID) ,M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin AS Admission_Date from MAINSTUDENTS M, OLDBOYS O where M.StuID=O.StuID
and M.StuID='20006'
union
select distinct(M.StuID),M.StudentName,M.FatherName,M.Street,M.City,M.D_Of_Admin AS Admission_Date from MAINSTUDENTS M, ACTIVESTUDENT A where M.StuID=A.StuID
and M.StuID='20006'

select Curr_Class,count(*) AS No_Of_Students from ACTIVESTUDENT group by(Curr_Class)

select * TEMPSTUDENTS


select * from OLRESULT
select * from ALRESULT

select * from OLRESULT O,MAINSTUDENTS M where O.StuID=M.StuID

select * from ACTIVESTUDENT A,MAINSTUDENTS M WHERE M.StuID=A.StuID and A.CURR_CLASS IN('11 R1','11 R2')
select * from BOOK B,COPY_OF_BOOK C where B.BookID=C.BookID and C.AccessNo='1000-1'
select * from OLRESULT
select * from BOOK
select * from COPY_OF_BOOK


select * from MAINSTUDENTS M , OLDBOYS O where M.StuID=O.StuID order by O.StuID asc

select * from OLDBOYS




select * from STAFFATTENDANCE  order by attDate,ComingTime asc

select * from SHORTLEAVES where LeaveType ='Short'

select * from SHORTLEAVES where month(dates)=month(GETDATE()) and LeaveType ='Short' and StaffID='5028'


select * from MEDICALLEAVES
select * from CASUALLEAVES
select * from STAFFLEAVES
select * from STAFFATTENDANCE


select * from SHORTLEAVES 

select StaffID,FullName from STAFF where StaffID NOT IN(
			select distinct(staffID) from STAFFATTENDANCE
			where day(attDate)=day(GETDATE()) and
			month(attDate)=month(GETDATE()) and
			year(attDate)=year(GETDATE()) and GoingTime IS NULL and
			NOT EXISTS
			(
			select StaffID from SHORTLEAVES 
			where day(dates)=day(GETDATE()) and
			month(dates)=month(GETDATE()) and
			year(dates)=year(GETDATE()) and ComingTime IS NULL))




print GETDATE()

SELECT CAST(GETDATE() AS datetime)

select * from SHORTLEAVES where CAST(Dates AS smalldatetime)=CAST(Dates AS smalldatetime)


select * from LENDING
select distinct(LE.MemID),L.MemName,L.Status,B.Title,LE.BorrowDate,LE.DueDate

select distinct(LE.MemID),L.MemName,L.Status,B.Title,LE.BorrowDate,LE.DueDate from LENDING LE,LIBRARYMEMBER L,COPY_OF_BOOK C ,BOOK B where C.BookID=B.BookID and
C.AccessNo=LE.AccessNo and L.MemID=LE.MemID and LE.Duedate<getDate()

print getdate()

select * from TEMPSTUDENTS


select * from COPY_OF_BOOK C ,BOOK B where C.BookID=B.BookID

select * from YEARAVERAGE
select * from MAINSTUDENTS
select * from ACTIVESTUDENT
select M.StuID,M.StudentName,M.FatherName,M.D_Of_Admin,M.AdminGrade,Year6,Year7,Year8,Year9,Year10,Year11,Year12,Year13 from MAINSTUDENTS M,YEARAVERAGE Y where Y.StuID=M.StuID and Y.StuID='20001'
select M.StuID,M.StudentName,M.FatherName,Year6,Year7,Year8,Year9,Year10,Year11,Year12,Year13 from MAINSTUDENTS M,YEARAVERAGE Y, ACTIVESTUDENT A where Y.StuID=M.StuID and A.StuID=M.StuID
select M.StuID,M.StudentName,M.FatherName,M.D_Of_Admin,AdminGrade from MAINSTUDENTS M
select FullName from STAFF where StaffID='1001'



