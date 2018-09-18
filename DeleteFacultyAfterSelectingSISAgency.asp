<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
AgencyID= Request("AgencyID")
SupervisorID = Request("SupervisorID")
Set objRS = Server.CreateObject("ADODB.recordset") 
        deleteFaculty_query = "delete from Supervisor1 where SupervisorID ='"&SupervisorID&"' "  
        deleteNotes_query = "delete from SupervisorNotes1 where SupervisorID ='"&SupervisorID&"' " 
        Response.Write "strsql1: " & deleteFaculty_query& "" &SupervisorID 
        objRS.open deleteFaculty_query, conn1
     Response.Write "strsql1: " & deleteNotes_query& "" &SupervisorID 
        objRS.open deleteNotes_query, conn1
conn1.close

        Response.Redirect "ViewAgency.asp?AgencyID="&AgencyID&"&ErrMsg='Instructor successfully deleted'"
%>