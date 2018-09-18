<html xmlns="http://www.w3.org/1999/xhtml">
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
    <div id="companyname1" align="center">
    <p><img src="images/JACSW-logo.png" alt="" width="265" height="80" /></p>
</div>
<div id="companyname" align="center">
    <p>
        <h1 class="AwardHeading">Student Information System </h1>
    </p>
</div>
<div align="center">
<div id="content">
<div id="steps">
<%if Session("Username") = "test_f" or Session("Username") = "tmorri3" or Session("Username") = "apradh6" or Session("Username") = "carrasc1" or Session("Username") = "bc1972" or Session("Username") = "nrosal1" or Session("Username") = "ktboyd" or Session("Username") = "chanze1" or Session("Username") = "fisherj" or Session("Username") = "melka1" or Session("Username") = "tashamc" or Session("Username") = "ajohns5"  Then %>
<h4><a href="ShowFieldStudents.asp">Home</a> | <a href="logout.asp">Log Out</a> | <a href="ShowCurrentStudents.asp">MSW Current Students</a> | <a href="ShowAllAgency.asp">Agency</a></h4>
    <% End If %>
</div>
</div>
</div>
</body>
</html>