<!--#include file="DBconn.asp"-->
<%
user_name = Request("userName")
programType = Request("programType")
roleType = Request("roleType")

set rs=Server.CreateObject("ADODB.recordset")
insert_query="insert into Roles(userID,roleType,proType) values (" & "'" & user_name & "','" & roleType & "','" & programType & "')"

query="select * from Roles where userID='" & user_name & "' and roleType='"  & roleType & "' and proType ='" & programType & "'"
rs.Open query,conn
	if not rs.EOF  then 
		Response.Redirect "AddRole.asp?UN="& user_name & "&ErrMsg='Role Already exists.'"
	else
		rs.close
		rs.Open insert_query,conn	
        Response.Redirect "UserRoles.asp?UN="& user_name & "&ErrMsg='Role was successfully added to DB'"	
	End if

rs.close
conn.close
%>