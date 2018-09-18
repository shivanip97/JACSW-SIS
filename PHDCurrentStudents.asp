<% 
ErrMsg = Request("ErrMsg")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>SIS | CURRENT PHD STUDENTS</title>
<link rel="stylesheet" href="css/tabstyle.css" type="text/css" />
    <script type="text/javascript" src="jquery/jquery-1.9.0.js"></script>
    <script type="text/javascript" src="jquery/filtertable.js"></script>
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
<!--#include file="headerPHDCurrent.asp"-->
<div align="center"><form action="" method="post"> 
     <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
     <br/> <br/>
     <div id="search">
       <p><label for="filter">Filter</label> <input type="text" name="filter" value="" id="filter" /></p> 
      </div>
      <br /><br />
       <table id="studentsTable">
	<thead>
		<tr>
            <th align="center">UIN</th>
            <th align="center">First Name</th>
            <th align="center">Last Name</th>
            <th align="center">Maiden Name</th>
            <th align="center">Email</th>
            <th align="center">Program Type</th>
             </tr>
	</thead>
    <tbody>
     <%
					set rs=Server.CreateObject("ADODB.recordset")
					course_query = "select * from CurrentPHDStudents order by LastName"
         rs.Open course_query,conn 
                    If rs.EOF Then
             Response.write("No Students Found")
         Else
             Do While NOT rs.Eof
                 uin = rs("UIN")
           %>
		<tr>
            <td align="center"><div> <% Response.write rs("UIN") %></div></td>
			<td align="center"><div class="edit" id="<%= uin%> FirstName"><% Response.write rs("FirstName") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> LastName"> <% Response.write rs("LastName") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> MaidenName"> <% Response.write rs("MaidenName") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> email"> <% Response.Write rs("Email") %></div></td>
            <td align="center"><div> <% Response.Write rs("Type") %></div></td>
            <td><button type="submit" name="Button1" onclick="this.form.action='ShowPHDStudentRecords.asp'; this.forms.submit();" id="Button1" value='<% Response.write rs("UIN") %>'>View Records</button></td>
         </tr>
		 <% rs.MoveNext   
        Loop End If %>
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