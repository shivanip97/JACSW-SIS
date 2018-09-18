<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%

AddressId = Request("AddressId")
Set objRS = Server.CreateObject("ADODB.recordset") 
        deleteAddress_query = "delete from AgencyAddress where AddressId ='"&AddressId&"' "  
   
        Response.Write "strsql1: " & deleteAddress_query& "" &AddressId 
        objRS.open deleteAddress_query, conn1
conn1.close

        Response.Redirect "ShowAllAgency.asp" 
%>