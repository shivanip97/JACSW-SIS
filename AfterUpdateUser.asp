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
else if access =1 then
role = "Admin"
end if
end if
lastlogin = Date() & " - " & Time()
Response.Write(lastlogin)&"<br/>"
Response.Write(access)

set rs=Server.CreateObject("ADODB.recordset")
update_query="update Users set Password='"& password & "', Name='" & name & "', Email='"& email & "', AccessLevel='"& access & "', Role='"& role & "' where Username='" & user_name & "'"
		rs.Open update_query,conn
		Response.Redirect "ShowUsers.asp?ErrMsg='User Information successfully updated'"
		
rs.close
conn.close
%>