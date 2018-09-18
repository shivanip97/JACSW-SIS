<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
user_name = Request("username")
password = md5(Request("password"))
name = Request("firstname")
email = Request("email")
access = int(Request("access"))
if access = 0 then
role = "User"
else if access = 1 then
role = "Admin"
end if
end if
lastlogin = "Yet to Login"

set rs=Server.CreateObject("ADODB.recordset")
insert_query="insert into Users(Username, Password, Name, Email, AccessLevel, Role, LastLogin) values (" & "'" & user_name & "','" & password & "','" & name & "','" & email & "','" & access & "','" & role & "','" & lastlogin & "')"

query="select * from Users where Username='" & user_name& "'"
rs.Open query,conn
	if not rs.EOF  then 
		Response.Redirect "AddUser.asp?ErrMsg='Username already exists, please select a different Username'"
	else
		rs.close
		rs.Open insert_query,conn	
        Response.Redirect "ShowUsers.asp?ErrMsg='User was successfully added to DB'"	
	End if

rs.close
conn.close
%>
