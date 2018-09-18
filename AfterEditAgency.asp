<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%

FieldInstructor = Request("fieldInst")
Term=Request("term")
FieldTypeYear=Request("fty")
Agency = Replace(Request("agency"),"'","''")
UID = Request("agencyID")
    UIN = Request.QueryString("UIN")



update_query="update AddAgency1 set Term='"&Term&"',FieldTypeYear='"&FieldTypeYear&"',FieldInstructor='"&FieldInstructor&"' where UID='"&UID&"'"
    updateField_query="update Field1 set Term='"&Term&"',FieldTypeYear='"&FieldTypeYear&"',FieldInstructor='"&FieldInstructor&"' where UIN='"&UIN&"' and Agency = '"&Agency&"'"
	
    Set objRS = Server.CreateObject("ADODB.recordset") 
   objRS.open update_query,conn
    Set objRS1 = Server.CreateObject("ADODB.recordset") 
   objRS1.open updateField_query,conn
conn.close
conn1.close

        Response.Redirect "ShowAgency.asp?UIN="&UIN&"&ErrMsg='Agency Updated to DB'"
%>
