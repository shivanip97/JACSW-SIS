<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
id = Request.form("id")
value = Request.form("value")

WordArray = Split(id," ")

set rs=Server.CreateObject("ADODB.recordset")
update_query="update Applicants set "&WordArray(1)&"='"& value & "' where UIN='" & WordArray(0) & "'"
		rs.Open update_query,conn

Response.Redirect "PassValue.asp?value="&value

rs.close
conn.close

 %>