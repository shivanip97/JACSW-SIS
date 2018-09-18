<% 
ErrMsg = Request("ErrMsg")
user_name = Request("UN")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="md5.asp"-->
<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>SIS | SHOW ROLES</title>
<link rel="stylesheet" href="css/tabstyle.css" type="text/css" />
	<script type="text/javascript" src="jquery/jquery-1.8.2.min.js"></script>
    <script type="text/javascript" src="jquery/jquery-ui-1.7.2.custom.min.js"></script>
    <script type="text/javascript" src="jquery/jquery.chromatable.js"></script>
	<script type="text/javascript">
	    $(document).ready(function () {
	});
 	</script>
 </head>
<style type="text/css">
		table {
			text-align: left;
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


<body bgcolor="#f2f2f2">
<!--#include file="header.asp"-->
<!--#include file="headerAdmin.asp"-->
<div align="center"><form action="" method="post"> 
<br/>
<a href='ShowUsers.asp'>Back to Show Users</a> 
<br/> <br/>
     <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
     <br/> <br/>
       <table id="Roles" cellspacing="1" align="center" width="90%">
	<thead>
		<tr>
            <th align="center">User Name</th>
            <th align="center">Role Type</th>
            <th align="center">Program Type</th>
            <th align="center"><button type="submit" name="Button1" onclick="this.form.action='AddRole.asp?UN=<% Response.Write(user_name) %>'; this.forms.submit();" id="Button1" value='<% Response.Write(user_name) %>'>Add New Role</button></th>
		</tr>
	</thead>
           <%
					set rs=Server.CreateObject("ADODB.recordset")
					course_query="select * from Roles where userID like '"& user_name & "'"
					rs.Open course_query,conn 
                    If rs.EOF Then
                    Response.write("No Roles Found.")
                    Else
                    Do While NOT rs.Eof 
                  
           %>
	<tbody>
		<tr>
			<td align="center"><% Response.write rs("userID") %></td>
            <td align="center"><% Response.write rs("roletype") %></td>
            <td align="center"><% Response.write rs("proType") %></td>
            <td><button type="submit" name="Button2" onclick="this.form.action='RemoveRole.asp?UN=<% Response.write rs("userID") %>'; this.forms.submit();" id="Button2" value='<% Response.write rs("roleID") %>'>Remove Role</button></td>
           </tr>
		 <% rs.MoveNext   
        Loop End If %>
	</tbody>
</table>
<% rs.Close
                Set rs=Nothing
                conn.Close
                Set conn=Nothing
                %>
</form> 
<!--#include file="footer.asp"-->
</div>
<p>&nbsp;</p>

</body>
</html>