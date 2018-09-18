<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UIN = Request("UIN")
set rs=Server.CreateObject("ADODB.recordset")
query="select * from Field where UIN like '"& UIN & "'"
rs.Open query,conn

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>SIS | Field</title>
<link rel="stylesheet" href="css/tabstyle.css" type="text/css" />
    <script type="text/javascript" src="jquery/jquery-1.9.0.js"></script>
    <script type="text/javascript" src="jquery/filtertable.js"></script>
    <script type="text/javascript" src="jquery/jquery.jeditable.mini.js"></script>
	<script type="text/javascript">
	    $(document).ready(function () {
	       
	    });
 	</script>
    <style type="text/css">
		table {
			text-align: left;
			font-size: 12px;
			font-family: verdana;
			background: #c0c0c0;
		}
 
		table thead tr,
		table tfoot tr {
			background: #c0c0c0;
			height:50px;
		}
 
		table tbody tr {
			background: #f0f0f0;
		}
 
		td, th {
			border: 1px solid white;
		}
	form button {
	border:none;
	outline:none;
    -moz-border-radius: 10px;
    -webkit-border-radius: 10px;
    border-radius: 10px;
    color: #ffffff;
    display: block;
    cursor:pointer;
    margin: 0px auto;
    clear:both;
    padding: 5px 15px;
    text-shadow: 0 1px 1px #777;
    font-weight:bold;
    font-family:"Century Gothic", Helvetica, sans-serif;
    font-size:20px;
    -moz-box-shadow:0px 0px 3px #aaa;
    -webkit-box-shadow:0px 0px 3px #aaa;
    box-shadow:0px 0px 3px #aaa;
    background:#4797ED;
}
    form button:hover {
    background:#d8d8d8;
    color:#666;
    text-shadow:1px 1px 1px #fff;
}
	</style>
 </head>
<body bgcolor="#f2f2f2">
<!--#include file="header.asp"-->
<!--#include file="headerField.asp"-->
<div align="center">

    <form action="" method="post"> 
                <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
     <br/><br/>
           
     <p><label for="UIN">UIN: </label><strong><font color="#000000"><% Response.Write(UIN) %> </font></strong></p>

     <br/>
        <%
                    set rs=Server.CreateObject("ADODB.recordset")
					course_query1="select FirstName, LastName from CurrentStudents where UIN = '"& UIN & "'"
					rs.Open course_query1,conn
                    If rs.EOF Then
                   Else
            %>
                    <p><label for="Name:">Name: </label><strong><font color="#000000"><% Response.Write rs("FirstName") %> &nbsp <% Response.Write rs("LastName") %> </font></strong></p>
                    
                  <% End If %>
        
 <br/>
     <div id="search">
       <p><label for="filter">Filter</label> <input type="text" name="filter" value="" id="filter" /></p> 
      </div>
      <br /><br />
        <%
                    set rs=Server.CreateObject("ADODB.recordset")
					course_query="select * from Field where UIN like '"& UIN & "'"
					rs.Open course_query,conn
                    If rs.EOF Then
                   
            %>
                    <button type="submit" name="Submit" onclick="this.form.action='AddField.asp?UIN=' + this.value; this.forms.submit();" value='<% Response.write(UIN) %>'>Add New Field</button><br /><br />
                  <% End If %>
        

    <table id="studentsTable">
	<thead>

		<tr>
            <th align="center">Field Type</th>
            <th align="center">Field Type Year</th>
            <th align="center">Working Liasion Foundation</th>
            <th align="center">Field Liasion Foundation</th>
            <th align="center">Foundation Term</th>
            <th align="center">Working Liasion Concentration</th>
            <th align="center">Field Liasion Concentration</th>
            <th align="center">Concentration Term</th>
            
            

			        
		</tr>
	</thead>
    <tbody>
	     <%
                    set rs=Server.CreateObject("ADODB.recordset")
					course_query="select * from Field where UIN like '"& UIN & "'"
					rs.Open course_query,conn 
                    If rs.EOF Then
                    Response.write("No Fields Found.")
                    Else
                   ' Do While NOT rs.Eof
                    uin = rs("UIN")

      %>
        <tr>
			<td align="center"><% Response.write rs("FieldType") %></td>
            <td align="center"><div class="edit" id="<%= uin%> FieldTypeYear"> <% Response.write rs("FieldTypeYear") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> WorkingLiasion Foundation"> <% Response.write rs("WorkingLiasionFoundation") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> FacultyLiasionFoundation"> <% Response.write rs("FacultyLiasionFoundation") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> WorkingLiasion Foundation Term"> <% Response.write rs("WorkingLiasionFoundationTerm") %></div></td>
            
            <td align="center"><div class="edit" id="<%= uin%> WorkingLiasion Concentration"> <% Response.write rs("WorkingLiasionConcentration") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> FacultyLiasionConcentration"> <% Response.write rs("FacultyLiasionConcentration") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> WorkingLiasion Concentration Term"> <% Response.write rs("WorkingLiasionConcentrationTerm") %></div></td>
            
            <td><button type="submit" name="Button1" onclick="this.form.action='ViewField.asp?UIN=<%Response.Write (UIN) %>'; this.forms.submit();" id="Button1" value='<% Response.write (UIN) %>'>Edit Field</button></td>
            <td><button type="submit" name="Button2" onclick="this.form.action='ShowAgency.asp?UIN=<%Response.Write (UIN) %>'; this.forms.submit();" id="Button2" value='<% Response.write (UIN) %>'>Show Placement</button></td> 
            

        </tr>
		  
         <% End If %>
	</tbody>
    <% rs.Close
                Set rs=Nothing
                conn.Close
                Set conn=Nothing
                %>
</table>
</form> 
<!--#include file="footer.asp"-->
</div>
<!-- overlayed element -->
<div class="apple_overlay" id="overlay">
  <!-- the external content is loaded inside this tag -->
  <div class="contentWrap"></div>
</div>
<p>&nbsp;</p>

</body>
</html>