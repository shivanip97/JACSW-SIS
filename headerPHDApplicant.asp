﻿<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file="Login_Check.asp"-->
<head>
<style type="text/css">
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
</head>
<body>
<div align="center">
<div id="content">
<div id="steps">

<%if Session("Username") = "test_ap" or Session("Username") = "cstoakl" or Session("Username") = "tmorri3"Then %>
<h4><a href="PHDlogin.asp">Home</a>| <a href="logout.asp">Log Out</a>| <a href="PhDApplicationReports.asp?ID=220158">PhD Reports</a></h4>
<%Else %>
<h4><a href="MSWApplicationLogin.asp">Home</a> | <a href="logout.asp">Log Out</a></h4>
<%End If %>
</div>
</div>
</div>
</body>
</html>