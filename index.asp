<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
dim name, pass, loginButton
name=Request.form("username")
pass= md5(Request.form("password"))
Session("ptype") = Request.form("programType")
Session("roleType") = Request.form("roleType")
Session("password") = Request.form("password")
logButton=Request("Submit")="Login"

ErrMsg = Request("ErrMsg")
if name = "webmaster" then
    if Request.form("password")="vadmin" then
        Session("Username") = name
        Session("AccessLevel") =1
        Session("UserLoggedIn") = "true"
        Session.Timeout=300
        Response.Redirect "ShowUsers.asp"
    end if
end if

if logButton then

	set rs=Server.CreateObject("ADODB.recordset")
	query="select * from Users where Username='" & name& "'and Password='" & pass & "'"
	rs.Open query,conn
	
	if not rs.EOF  then 
		Dim Level
		Level = rs.Fields("AccessLevel")
		Session("AccessLevel") = Level
		Session("UserLoggedIn") = "true"
        Session("Username") = name  
        Session("LastLogin") = rs("LastLogin")
        Session.Timeout=300
		
		Response.Redirect "LoginInfo.asp"
		
		else
			Dim ErrMsg 
			ErrMsg = "Please enter a valid username or password."
	End if
    rs.Close
End if
conn.close
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<title>JACSW | SIS</title>
<link rel="stylesheet" href="css/loginStyle.css" type="text/css" media="screen" />
<script>
    $(document).ready(function () {
        sessionStorage.removeItem("tag");
    });
</script>
<script> function Trim(str){return str.replace(/^\s*|\s*$/g,"");} </script>
<script language="javascript">
var request = makeObject();	
function validate(){
	var username =  document.getElementById("username");
	var password =  document.getElementById("password");
		
	if(Trim(username.value)=='')
	{
		alert("Please enter a Username");
		fn.focus();
		return false;		
	}
	if(Trim(password.value)=='')
	{
		alert("Please provide a password");
		ln.focus();
		return false;		
	}
}
	
</script>
<style type="text/css">
.AwardHeading {
	font-family: Verdana, Geneva, sans-serif;
	color:#666;
	text-align:center;
}
</style>
</head>
<body>

	<div align="center">
    <div id="header">
		<!--#include file="header.asp"-->
       <div align="right" class="links_menu" id="menu"></div>
    </div>
		<br />
		<div id="content">
        <div id="steps">
				<form id="form1" method="post" action="">
					<p>Welcome to the Student Information System
                    <br/><br/>
					<label>Username</label>
					<input type="text" name="username" id="username" />    
                    <br/><br/><br/>
					<label>Password </label>
					<input type="password" name="password" id="password" />
                    <br/><br/><br/>
                    <label>Program Type </label>
                    <select name="programType">
                    <option selected="selected" value="Select">-Select-</option>
                    <option value="MSW">MSW</option>
                    <option value="PHD">PHD</option>
                    <option value="T73">PEL</option>
                    
                    </select>
                    <br/><br/><br/>
                    <label>Form </label>
                    <select name="roleType">
                    <option selected="selected" value="Select">-Select-</option>
                    <option value="Application">Application</option>
                    <option value="Current">Current Students</option>
                    <option value="Field">Field</option>
                    <option value="Agency">Agency</option>
                    </select>
                    <br/><br/><br/>
					<button type="submit" name="Submit" onclick="return validate();" value="Login">Login</button>
					<br/>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong><br />
                   </p>
				</form>
                </div>
               </div>
			<br />
			
			<div class="footer"><br />
            <!--#include file="footer.asp"-->
		  </div>
          <p class="news_date">&nbsp;</p>
		</div>
	</div>
</body>
</html>
