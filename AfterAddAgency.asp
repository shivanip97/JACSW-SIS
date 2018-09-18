<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
Term = Request("term")
FieldTypeYear = Request("fty")
UIN = Request("Submit")

AgencyID = Request("agency")
Session("AgencyID")=AgencyID
    set rs10=Server.CreateObject("ADODB.recordset")
					course_query="select a.AgencyID,a.Agency from Agency1 a where a.AgencyID = '"& AgencyID & "' "
					rs10.Open course_query,conn1 
Agency = Replace(rs10.Fields(1), "'", "''")

set rs4=Server.CreateObject("ADODB.recordset")
					course_query="select distinct LastName,FirstName,FacultyLiasionConcentration, WorkingLiasionConcentration,FacultyLiasionFoundation, WorkingLiasionFoundation,WorkingLiasionConcentrationTerm, WorkingLiasionFoundationTerm from Field1 where UIN = '"& UIN & "' "
					rs4.Open course_query,conn 
    LastName = rs4.Fields(0)
    FirstName = rs4.Fields(1)
    FacultyLiasionConcentration = rs4.Fields(2) 
    WorkingLiasionConcentration = rs4.Fields(3)
    FacultyLiasionFoundation = rs4.Fields(4) 
    WorkingLiasionFoundation = rs4.Fields(5)
    WorkingLiasionConcentrationTerm =rs4.Fields(6) 
    WorkingLiasionFoundationTerm =rs4.Fields(7)

set rs=Server.CreateObject("ADODB.recordset")

query="select Agency from Field1 where UIN = '" &UIN& "' and Agency ='" &Agency& "'"
rs.Open query,conn
	if not rs.EOF  then 
    Response.write (Agency)
		Response.Redirect "AddAgency.asp?UIN="&UIN&"&ErrMsg='Agency already exists, please select a different Agency'"
	else
		rs.close
    end if

    set rs7=Server.CreateObject("ADODB.recordset")
    query2="select Agency from Field1 where UIN='" &UIN& "'"
    rs7.Open query2,conn
    AgencyName=rs7.Fields(0)
    if IsNull(AgencyName)=TRUE then
    set rs6=Server.CreateObject("ADODB.recordset")	
    updatefield_query="update Field1 set Term= '"&Term&"',AgencyID='"&AgencyID&"',FieldTypeYear='"&FieldTypeYear&"',Agency='"&Agency&"' where UIN = '" &UIN& "'"
    rs6.Open updatefield_query,conn
    else
    set rs2=Server.CreateObject("ADODB.recordset")	
    insertfield_query="insert into Field1(Term,FieldTypeYear,Agency,AgencyID,UIN,LastName,FirstName,FacultyLiasionConcentration, WorkingLiasionConcentration,FacultyLiasionFoundation, WorkingLiasionFoundation,WorkingLiasionConcentrationTerm, WorkingLiasionFoundationTerm) values ('"&Term&"','"&FieldTypeYear&"','"&Agency&"','"&AgencyID&"','" &UIN& "','"&LastName&"', '"&FirstName&"','"&FacultyLiasionConcentration&"', '"&WorkingLiasionConcentration&"','"&FacultyLiasionFoundation&"', '"&WorkingLiasionFoundation&"','"&WorkingLiasionConcentrationTerm&"', '"&WorkingLiasionFoundationTerm&"') "
    rs2.Open insertfield_query,conn
	End if
     Response.Redirect "AddInstructorAgency.asp?UIN="&UIN&"&ErrMsg='Agency was successfully added to DB'"

rs.close
    
    rs2.close
    rs3.close
    rs4.close
    rs5.close
    rs6.close
    rs7.close
    rs8.close
conn.close
conn1.close
%>
