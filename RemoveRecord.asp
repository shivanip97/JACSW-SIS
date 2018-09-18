<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<%
UIN = Session("UIN")
id = Request("id")
set rs=Server.CreateObject("ADODB.recordset")
delete_query="delete from CurrentStudent where UIN='" & UIN & "' and ID ='"&id&"'"
rs.Open delete_query, conn
conn.close
Response.Redirect "ShowCurrentStudents.asp?ErrMsg='Record was successfully removed from DB'"
%>