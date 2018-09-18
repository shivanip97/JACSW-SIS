<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
UIN = Request("uin")
Fname = Request("fname")
Lname = Request("lname")
FieldType = Request("fieldtype")
FieldTypeYear = Request("fieldtypeyear")
FacultyLiasionFoundation=Request("flf")
FacultyLiasionConcentration = Request("flc")
WorkingLiasionConcentration = Request("wlc")
WorkingLiasionFoundation = Request("wlf")
InfoSent = Request("infoSent")
WorkingLiasionConcentrationTerm = Request("wlct")
WorkingLiasionFoundationTerm = Request("wlft")
DateCommentsEntered = Request("dce")
Comments = Request("comments")

set rs=Server.CreateObject("ADODB.recordset")
insert_query="insert into Field1(FieldType,FacultyLiasionFoundation,FacultyLiasionConcentration,WorkingLiasionConcentration,WorkingLiasionFoundation,WorkingLiasionConcentrationTerm,WorkingLiasionFoundationTerm,InfoSent,DateCommentsEntered,Comments,UIN,FirstName,LastName,FieldTypeYear) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

Set objCommand = Server.CreateObject("ADODB.Command") 
objCommand.ActiveConnection = conn
objCommand.CommandText = insert_query
objCommand.Parameters(0).value = FieldType
objCommand.Parameters(1).value = FacultyLiasionFoundation
objCommand.Parameters(2).value = FacultyLiasionConcentration
objCommand.Parameters(3).value = WorkingLiasionConcentration
objCommand.Parameters(4).value = WorkingLiasionFoundation
objCommand.Parameters(5).value = WorkingLiasionConcentrationTerm
objCommand.Parameters(6).value = WorkingLiasionFoundationTerm
objCommand.Parameters(7).value = InfoSent
objCommand.Parameters(8).value = DateCommentsEntered
objCommand.Parameters(9).value = Comments
objCommand.Parameters(10).value = UIN
objCommand.Parameters(11).value = Fname
objCommand.Parameters(12).value = Lname
objCommand.Parameters(13).value = fieldtypeyear

Set objRS = objCommand.Execute()
        Response.Redirect "ViewFieldNew.asp?UIN="&UIN & "&ErrMsg='User was successfully added to DB'"	
rs.close
conn.close
%>
