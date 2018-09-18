<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<%
UN = Request("UN")
roleID = Request("Button2")
set rs=Server.CreateObject("ADODB.recordset")
delete_query="delete from Roles where roleID='" & roleID & "'"
rs.Open delete_query, conn
conn.close
Response.Redirect "UserRoles.asp?UN="& UN &"&ErrMsg='Role was successfully Removed.'"
%>