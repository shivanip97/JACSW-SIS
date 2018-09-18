<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%

set rs=Server.CreateObject("ADODB.recordset")
query="select Max(AddressId) as max from AgencyAddress"
rs.Open query,conn1
Add = rs("max") + 1
AddressId = Add
AgencyID = Request("AgencyID")
AddressL1 = Request("AddressL1")
AddressL2 = Request("AddressL2")
City = Request("City")
State = Request("State")
Zip = Request("Zip")

Set objRS = Server.CreateObject("ADODB.recordset") 
     Response.Write (AddressId)
insertaddress_query="insert into AgencyAddress (AgencyID,AddressL1,AddressL2,City,State,Zip,AddressId,PrimaryAddress,MailingAddress) values ('"&AgencyID&"','"&AddressL1&"', '"&AddressL2&"', '"&City&"', '"&State&"', '"&Zip&"','"&AddressId&"','1','1')"
        Response.Write "strsql1: " & insertaddress_query 
        objRS.open insertaddress_query, conn1
conn1.close
    
        Response.Redirect "EditSISAgency.asp?AgencyID="&AgencyID&"&ErrMsg='Student Information Updated to DB'"
%>