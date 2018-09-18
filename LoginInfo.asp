<!--#include file="DBconn.asp"-->
<%
lastlogin = Date() & " - " & Time()
set rs=Server.CreateObject("ADODB.recordset")
update_query="update Users set LastLogin='"&lastlogin & "' where Username='" & Session("Username") & "'"
rs.Open update_query,conn
if Session("password") = "change_me" then
    Response.Redirect "ChangePassword.asp"
End If
If Session("AccessLevel") = 0 Then
        	Response.Redirect "ProgramTypeCheck.asp"
        Else If Session("AccessLevel") = 1 Then
        	Response.Redirect "ShowUsers.asp"
        End if
End if
	conn.close	
%>