<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
uin = Request("uin")
LastName = Request("lastname")
FirstName = Request("firstname")
MiddleName = Request("middlename")
MaidenName = Request("maidenname")
Gender = Request("gender")

CurrentAddress1 = Request("currentAddress1")
CurrentAddress2 = Request("currentAddress2")
CurrentCity = Request("currentcity")
CurrentState = Request("currentstate")
OAR_Application_Date = Request("appdate")
   
CurrentZipCode = Request("currentzip")
CurrentCountry = Request("currentcountry")
HomePhone = Request("homephone")
WorkPhone = Request("workphone")
InternationalPhoneNumber = Request("intphone")
EMail = Request("ugcollege")
DateOfBirth = Request("dob")
Comments = Request("comments")
UGCollege = Request("ugcollege")
UGGPA = Request("uggpa")
UGMajor = Request("ugmajor")
GradCollege = Request("graddegree")
GradMajor = Request("gradmajor")
GradGPA = Request("gradgpa")
GradDegree = Request("graddegree")

set rs=Server.CreateObject("ADODB.recordset")
insert_query="insert into Applicants(UIN,LastName,FirstName,MiddleName,Gender,CurrentAddress1,CurrentAddress2,CurrentCity,CurrentState,CurrentZipCode,CurrentCountry,HomePhone,WorkPhone,InternationalPhoneNumber,EMail,DateOfBirth,Comments,UGCollege,UGGPA,UGMajor,GradCollege,GradMajor,GradGPA,GradDegree,MaidenName,OAR_Application_Date) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

query="select * from Applicants where UIN='" & uin& "'"
rs.Open query,conn
	if not rs.EOF  then 
		Response.Redirect "AddStudent.asp?ErrMsg='UIN already exists, please select a different UIN'"
	else
		rs.close
		Set objCommand = Server.CreateObject("ADODB.Command") 
	                objCommand.ActiveConnection = conn
	                objCommand.CommandText = insert_query
	                objCommand.Parameters(0).value = uin
	                objCommand.Parameters(1).value = LastName
	                objCommand.Parameters(2).value = FirstName
objCommand.Parameters(3).value = MiddleName
objCommand.Parameters(4).value = Gender
objCommand.Parameters(5).value = CurrentAddress1
objCommand.Parameters(6).value = CurrentAddress2
objCommand.Parameters(7).value = CurrentCity
objCommand.Parameters(8).value = CurrentState
objCommand.Parameters(9).value = CurrentZipCode
objCommand.Parameters(10).value = CurrentCountry
objCommand.Parameters(11).value = HomePhone
objCommand.Parameters(12).value = WorkPhone
objCommand.Parameters(13).value = InternationalPhoneNumber
objCommand.Parameters(14).value = EMail
objCommand.Parameters(15).value = DateOfBirth
objCommand.Parameters(16).value = Comments
objCommand.Parameters(17).value = UGCollege
objCommand.Parameters(18).value = UGGPA
objCommand.Parameters(19).value = UGMajor
objCommand.Parameters(20).value = GradCollege
objCommand.Parameters(21).value = GradMajor
objCommand.Parameters(22).value = GradGPA
objCommand.Parameters(23).value = GradDegree
objCommand.Parameters(24).value = MaidenName
objCommand.Parameters(25).value = OAR_Application_Date
Set objRS = objCommand.Execute()
        Response.Redirect "ShowStudents.asp?ErrMsg='Student was successfully added to DB'"	
	End if

rs.close
conn.close
%>
