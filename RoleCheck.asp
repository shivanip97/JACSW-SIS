<!--#include file="DBconn.asp"-->
<%
        If Session("roleType") = "Application" Then
        	Response.Redirect "MastersApplication.asp"
        Else If Session("roleType") = "Field" Then
        	Response.Redirect "MastersField.asp"
        Else If Session("roleType") = "Current" Then
        	Response.Redirect "MastersCurrent.asp"
        End if
        end if
        end if
%>