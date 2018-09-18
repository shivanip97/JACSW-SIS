<% 
ErrMsg = Request("ErrMsg")
UIN = Request.QueryString("UIN")

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
	        $('.edit').editable('UpdateAgency.asp');
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

<div align="center">
           <form action="" method="post"> 
                <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
     <br/><br/>
                    <p><label for="UIN">UIN: </label><strong><font color="#000000"><% Response.write(UIN) %></font></strong></p>

     <br/> 
                <%
                    set rs1=Server.CreateObject("ADODB.recordset")
					course_query1="select FirstName, LastName from CurrentStudents where UIN = '"& UIN & "'"
					rs1.Open course_query1,conn
                    If rs1.EOF Then
                   Else
            %>
                    <p><label for="Name:">Name: </label><strong><font color="#000000"><% Response.Write rs1("FirstName") %> &nbsp <% Response.Write rs1("LastName") %> </font></strong></p>
                    
                  <% End If %>
        
 
     <br/> <br/>
     
               <a href="ViewFieldNew.asp?UIN=<%Response.Write (UIN) %>">Back to Show Field</a> 
      <br /><br />
     <div id="search">
       <label for="filter">Filter</label> <input type="text" name="filter" value="" id="filter" />
      </div>
    <br />
       <table id="agencyTable">
	<thead>
    
		<tr>
            <th align="center">Agency</th>
            
            <th align="center">Term</th>
            <th align="center">Field Type Year</th>
            <th align="center">Field Instructor</th>
            <th align="center"><button type="submit" name="Button" onclick ="this.form.action='AddAgency.asp?UIN=<%Response.Write (UIN) %>'; this.forms.submit();" id="Button" value='<% Response.Write (UIN) %>'>Add Placement</button> </th>
		</tr>
	</thead>
    <tbody>
     <%
					set rs=Server.CreateObject("ADODB.recordset")
					course_query="select f.Term as Term,f.FieldTypeYear as FieldTypeYear,f.Agency as Agency, f.FieldInstructor as FieldInstructor, f.UID as UID FROM Field1 f where f.UIN = '"& UIN & "' and Agency IS NOT NULL order by Term"
					rs.Open course_query,conn 
                    If rs.EOF Then
                    Response.write("No Agencies Found.")
                    Else
                    Do While NOT rs.Eof  
                    uid = rs("UID")
           %>
		<tr>
            <td align="center"><div id="<%= uin%> Agency"> <% Response.write rs("Agency") %></div></td>
			
            <td align="center"><div class="edit" id="<%= uin%> Term"> <% Response.write rs("Term") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> FieldTypeYear"> <% Response.write rs("FieldTypeYear") %></div></td>
            <td align="center"><div id="<%= uin%> FieldInstructor"> <% Response.write rs("FieldInstructor") %></div></td>
            <td><button type="submit" name="Button1" onclick="this.form.action='EditAgency.asp?UIN=<%Response.Write (UIN) %>'; this.forms.submit();" id="Button1" value='<% Response.write rs("uid") %>' ">Edit</button></td>
         </tr>
		 <%   rs.MoveNext 
             Loop
             End If %>
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