<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<%
user_name = Request("UN")
set rs=Server.CreateObject("ADODB.recordset")
delete_query="delete from Users where Username='" & user_name & "'"
rs.Open delete_query, conn
conn.close
Response.Redirect "ShowUsers.asp?ErrMsg='User was successfully removed from DB'"
%>