<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%

AgencyID = Request("AgencyID")
Set objRS = Server.CreateObject("ADODB.recordset") 
        deleteAgencyDataforOnline_query = "delete from AgencyDataForOnline1 where AgencyID ='"&AgencyID&"' " 

    Set objRS1 = Server.CreateObject("ADODB.recordset")
    deleteAgency_query = "delete from Agency1 where AgencyID ='"&AgencyID&"' "  

     Set objRS2 = Server.CreateObject("ADODB.recordset")
    deleteAgencyNotes_query = "delete from AgencyNotes1 where AgencyID ='"&AgencyID&"' "
    
       Set objRS3 = Server.CreateObject("ADODB.recordset")
    deleteAgencyAddress_query = "delete from AgencyAddress1 where AgencyID ='"&AgencyID&"' "
   
        Response.Write "strsql1: " & deleteAgency_query& "" &AgencyID 
    
    objRS.open deleteAgencyDataforOnline_query, conn1
    objRS1.open deleteAgency_query, conn1
    objRS2.open deleteAgencyNotes_query, conn1
    objRS3.open deleteAgencyAddress_query, conn1
conn1.close

        Response.Redirect "ShowAllAgency.asp" 
%>