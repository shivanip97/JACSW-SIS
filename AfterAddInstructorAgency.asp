<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%

FieldIns = Request("fieldInst")

agency = Replace(Request("agency"),"'","''")
UID = Request("uid1")
UIN = Request.QueryString("UIN")

    

    Set objRS = Server.CreateObject("ADODB.recordset") 
     
   update_query="update Field1 set FieldInstructor='"&FieldIns&"' where UID ='"&UID&"'"
   objRS.open update_query,conn
   

    
    
    
    
conn.close
conn1.close

        Response.Redirect "ShowAgency.asp?UIN="&UIN&"&ErrMsg='Agency Updated to DB'"
%>
