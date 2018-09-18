<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>SIS | View Students</title>
    <!--#include file="DBconn.asp"-->
    <!--#include file="Login_Check.asp"-->
    <!--#include file="header.asp"-->
    <% 
%>
    <link rel="stylesheet" href="css/tabstyle.css" type="text/css" media="screen" />
	<script type="text/javascript" src="jquery/jquery-1.8.2.min.js"></script>
    <script type="text/javascript" src="jquery/filtertable.js"></script>
    <script type="text/javascript" src="jquery/jquery-ui-1.8.2.custom.min.js"></script>
    <script type="text/javascript" src="jquery/sliding.form.js"></script>
    <script>
        $(document).ready(function () {
        });

        function getval(sel) {
            window.location = "https://socialwork.cc.uic.edu/SIS/ViewMSW.asp?ID=" + sel.value;
        }
    </script>
        <style type="text/css">
		table {
			text-align: center;
			font-size: 12px;
			font-family: verdana;
			background: #c0c0c0;
			table-layout:fixed;
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
			height:50px;
			overflow: hidden;
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

<body>
<div align="center">
<br />
    <a style="font-size:12pt;" href='PHDApplication.asp?ID=Fall2015'>Back to PHD Students</a>
    <form action="" method="post"> 
     <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
     <br/> <br/>
     
      <label>Program Type</label>
      <select name="program_type" id="program_type" onchange="getval(this);">
      <option value="<%=Request("ID") %>" selected> <%=Request("ID") %></option>
      <option value="All">All</option>
      <option value="FT">FT</option>
      <option value="PM">PM</option>
      <option value="Adv">Adv</option>
      <option value="TR">TR</option>
      <option value="MPH-Adv">MPH-Adv</option>
      <option value="MPH-FT">MPH-FT</option>
      <option value="MPH-PM">MPH-PM</option>
      </select> 
      <br /><br />
<div id="search">
       <p><label for="filter">Filter</label> <input type="text" name="filter" value="" id="filter" /></p> 
      </div>
      <br /><br />
    <div id="content">
        <h2><%=Request("ID")%> Students</h2><br />
        <form id="formElem" name="formElem" action="" method="post">
            <div id="students" align="center">
         <table id="studentsTable">
	<thead>
		<tr>
            <th align="center">UIN</th>
            <th align="center">First Name</th>
            <th align="center">Last Name</th>
            <th align="center">Maiden Name</th>
            <th align="center">Email</th>
         </tr>
	</thead>
             </div>
    <tbody>
     <%
                  Prog_Name= Request("ID")
                  
					set rs=Server.CreateObject("ADODB.recordset")
                    If ((StrComp(Prog_Name,"All"))= 0) Then
					course_query="select * from Applicants order by LastName"
                    Else
                    course_query="select * from Applicants where program_type like '" & Request("ID") & "' order by LastName"
                    End If
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
            <td align="center"><div class="edit" id="<%= uin%> email"> <% Response.write rs("email") %></div></td>
            <td><button type="submit" name="Button1" onclick="this.form.action='ViewStudentInfo.asp'; this.forms.submit();" id="Button1" value='<% Response.write rs("UIN") %>'>View Records</button></td>

         </tr>
		 <% rs.MoveNext   
        Loop End If %>
	</tbody>
    <% rs.Close
                Set rs=Nothing
                %>
</table>
</form>
        </div>
    <!--#include file="footer.asp"-->
</body>
</html>

