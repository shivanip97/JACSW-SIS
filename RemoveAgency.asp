<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<%
 UID = Request("Submit1")
 UIN = Request.QueryString("UIN")


    set rs2=Server.CreateObject("ADODB.recordset")
    selectfield_query="select * from Field1 where UIN ='" & UIN & "'"
    rs2.Open selectfield_query, conn, 1,1

    if rs2.RecordCount<>1 then

  set rs1=Server.CreateObject("ADODB.recordset")
    deletefield_query="delete from Field1 where UIN ='" & UIN & "' and UID ='" & UID & "'"
    rs1.Open deletefield_query, conn
    
    
    else
     set rs3=Server.CreateObject("ADODB.recordset")
    updatefield_query="update Field1 set Agency= null,FieldTypeYear= null,Term= null,FieldInstructor=null where UIN ='" & UIN & "' and UID ='" & UID & "'"
    rs3.Open updatefield_query, conn

    end if
    Response.Redirect "ShowAgency.asp?UIN="&UIN&"&ErrMsg='Agency was successfully removed from DB'"
    rs1.close
     rs2.close
    rs3.close
    conn.close
%>