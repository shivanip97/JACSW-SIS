<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%

set rs=Server.CreateObject("ADODB.recordset")
query="select Max(SupervisorID)+1 as MaxID from Supervisor1"
rs.Open query,conn1
Sup = rs("MaxID") 

AgencyID = Request("AgencyID")
SupervisorFullName =Replace(Request("SupervisorFullName"),"'","''") 
EmailAddress = Replace(Request("EmailAddress"),"'","''")
Phone = Request("SPhone")
CellPhone = Request("CellPhone")

SupervisorID = Sup

Set objRS = Server.CreateObject("ADODB.recordset") 
     Response.Write (Sup)
insertfaculty_query="insert into Supervisor1 (AgencyID,SupervisorFullName,EmailAddress,Phone,CellPhone,SupervisorID) values ('"&AgencyID&"','"&SupervisorFullName&"', '"&EmailAddress&"', '"&Phone&"', '"&CellPhone&"','"&SupervisorID&"')"
        Response.Write "strsql1: " & insertfaculty_query 
        objRS.open insertfaculty_query, conn1
conn1.close
    
        Response.Redirect "EditSISAgency.asp?AgencyID="&AgencyID&"&ErrMsg='Student Information Updated to DB'"
%>