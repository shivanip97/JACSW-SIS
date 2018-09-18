<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
user_name = Session("Username")
password = md5(Request("password"))
lastlogin = Date() & " - " & Time()
set rs=Server.CreateObject("ADODB.recordset")
update_query="update Users set Password='"& password & "' where Username='" & user_name & "'"
		rs.Open update_query,conn
		Response.Redirect "index.asp?ErrMsg='Your password has been changed successfully'"
rs.close
conn.close
%>
