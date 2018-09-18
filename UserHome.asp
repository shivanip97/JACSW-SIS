<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>SIS | HOME</title>
<link rel="stylesheet" href="css/tabstyle.css" type="text/css" />
	<script type="text/javascript" src="jquery/jquery-1.8.2.min.js"></script>
    <script type="text/javascript" src="jquery/jquery-ui-1.7.2.custom.min.js"></script>
 </head>
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
h5{
    color: #000066;
    font-size: 16px;
}

ul
{
list-style-type:none;
margin:0 auto;
padding:0 1px;
text-align:center;
overflow:hidden;
}
li
{
float:left;
}
a:link,a:visited
{
display:block;
width:120px;
font-size:12pt;
font-weight:normal;
color:#666;
background-color:#f4f4f4;
text-align:center;
padding:4px;
text-decoration:none;
text-transform:uppercase;
}
a:hover,a:active
{
background-color:#d8d8d8;
}
	</style>
<body bgcolor="#f2f2f2">
<!--#include file="header.asp"-->
<div align="center">
<div id="content">
<div id="steps">
<h4><a href="UserHome.asp">Home</a> | <a href="AddStudent.asp">Add New Student</a> | <a href="logout.asp">Log Out</a></h4>
</div>
</div>
</div>
<div align="center">
<form action="" method="post"> 
</br></br>
</br></br>
       <h5 align='center'>Welcome, <% Response.write Session("Username") %>. Your Last Login was on : <% Response.write Session("LastLogin") %> </h5></br>
                <h3>If you have any questions please e-mail Vivek at vvenka6@uic.edu</h3></br>
                </br></br>
      <button type="submit" name="Button1" onclick="this.form.action='ShowStudents.asp'; this.forms.submit();" id="Button2">Prospective Basic Information</button>
       </br></br>
      <button type="submit" name="Button1" onclick="this.form.action='ShowInfoStudents.asp'; this.forms.submit();" id="Button1">Prospective Additional Information</button>
       </br></br>
      <button type="submit" name="Button1" onclick="this.form.action='ShowCurrentStudents.asp'; this.forms.submit();" id="Button3">Current Information</button>
      </br></br>
      <button type="submit" name="Button1" onclick="this.form.action='ShowFieldStudents.asp'; this.forms.submit();" id="Button4">Field Information</button>

      </br></br>
</form> 
</br></br>

<!--#include file="footer.asp"-->
</div>
<p>&nbsp;</p>
</body>
</html>