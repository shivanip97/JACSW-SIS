<!--#include file="DBconn.asp"-->
<%
set rs=Server.CreateObject("ADODB.recordset")
query="select proType,roleType from Roles where userID='" & Session("Username") & "'and proType='" & Session("ptype") & "' and roleType='" & Session("roleType") & "'"
rs.Open query,conn
if not rs.EOF  then 
        If Session("ptype") = "MSW" Then
            If Session("roleType") = "Application" Then
        	    Response.Redirect "MSWApplicationLogin.asp"
            Else If Session("roleType") = "Field" Then
        	    Response.Redirect "ShowFieldStudents.asp"
            Else If Session("roleType") = "Current" Then
        	    Response.Redirect "ShowCurrentStudents.asp"
            Else If Session("roleType") = "Agency" Then
        	    Response.Redirect "ShowAllAgency.asp"
            End if
            end if
            end if
            end if
        Else If Session("ptype") = "PHD" Then
        	If Session("roleType") = "Application" Then
        	    Response.Redirect "PHDlogin.asp"
            Else If Session("roleType") = "Current" Then
        	    Response.Redirect "PHDCurrentStudents.asp"
            End if
            end if
        Else If Session("ptype") = "T73" Then
        	If Session("roleType") = "Application" Then
        	    Response.Redirect "T73Application.asp"
            Else If Session("roleType") = "Current" Then
        	    Response.Redirect "T73Current.asp"
            End if
            end if
        Else If Session("ptype") = "MPH" Then
        	If Session("roleType") = "Application" Then
        	    Response.Redirect "MPHApplication.asp"
            Else If Session("roleType") = "Current" Then
        	    Response.Redirect "MPHCurrent.asp"
            Else If Session("roleType") = "Field" Then
        	    Response.Redirect "MPHField.asp"
            End if
            end if
            end if
        End If
        End if
        End If
        End if
else
    Response.Redirect "index.asp?ErrMsg='Invalid Program or Role Type.'"
End if
	conn.close	
%>