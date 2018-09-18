<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<%
UIN = Request("UIN")
set rs=Server.CreateObject("ADODB.recordset")
delete_query="delete from PHDApplicants where UIN='" & UIN & "'"
rs.Open delete_query, conn
conn.close
Response.Redirect "PHDCurrentStudents.asp?ErrMsg='User was successfully removed from DB'"
%>