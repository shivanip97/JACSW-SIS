<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
UIN = Request("uin")

    FieldTypeYear = Request("fieldtypeyear")

FacultyLiasionFoundation=Request("flf")
FacultyLiasionConcentration = Request("flc")
WorkingLiasionConcentration = Request("wlc")
WorkingLiasionFoundation = Request("wlf")

WorkingLiasionConcentrationTerm = Request("wlct")
WorkingLiasionFoundationTerm = Request("wlft")

update_query="Update Field1 set FieldType='MSW', FacultyLiasionFoundation='"&FacultyLiasionFoundation&"',FacultyLiasionConcentration='"&FacultyLiasionConcentration&"',WorkingLiasionConcentration='"&WorkingLiasionConcentration&"',WorkingLiasionFoundation='"&WorkingLiasionFoundation&"',WorkingLiasionConcentrationTerm='"&WorkingLiasionConcentrationTerm&"',WorkingLiasionFoundationTerm='"&WorkingLiasionFoundationTerm&"' where UIN like '"&UIN&"' "

    Set objRS = Server.CreateObject("ADODB.recordset") 
     Response.Write "strsql1: " & update_query 
   objRS.open update_query, conn
   
  
conn.close


        Response.Redirect "ViewFieldNew.asp?UIN="&UIN&"&ErrMsg='Field was successfully Updated to DB'"	

objRS.close
conn.close
%>
