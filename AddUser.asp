<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>FAAR | Add New User</title>
<link rel="stylesheet" href="css/newUserStyle.css" type="text/css" media="screen" />
 <script>     function Trim(str) { return str.replace(/^\s*|\s*$/g, ""); } </script>
<script language="javascript">
    var request = makeObject();
    function validate() {
        var username = document.getElementById("userName");
        var fname = document.getElementById("firstname");
        var lname = document.getElementById("lastname");
        var email = document.getElementById("email");
        var access = document.getElementById("access");

        if (Trim(username.value) == '') {
            alert("Please enter a Username");
            username.focus();
            return false;
        }
        if (Trim(fname.value) == '') {
            alert("Please provide a Name");
            fname.focus();
            return false;
        }
        if (Trim(email.value) == '') {
            alert("Please provide an Email");
            email.focus();
            return false;
        }
        if (Trim(access.value) == '2') {
            alert("Please select the User Access Level");
            access.focus();
            return false;
        }
    }
	
</script>
</head>
<body>
    <div id="content" align=center>
            <div id="steps">
				<form id="form1" method="post" action="AfterAddUser.asp">
					<h3>Add New User</h3>
                     <br/>
                    <a href=ShowUsers.asp>Back to Show Users</a> 
                    <br/> <br/>
                    <p>
                    <label>Add User Form</label>
                    <br/><br/><br/>
                     <label>User Name</label>
					<input type="text" name="userName" id="userName"/> 
                    <br/><br/><br/>
                    <label>Password</label>
					<input type="text" name="password" id="password" value="change_me" readonly=true/> 
                    <br/><br/><br/>
                    <label>Name</label>
					<input type="text" name="firstname" id="firstname"/>    
                    <br/><br/><br/>
                    <label>Email</label>
					<input type="text" name="email" id="email" />
                    <br/><br/><br/>
                     <label>Access Level</label>
   	                <select name="access" id="access">
         			<option value="2">-- Select --</option>
  					<option value="1">Admin</option>
					<option value="0">User</option>
				    </select>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/><br/><br/>
					<button type="submit" name="Submit" onclick="return validate();">Add User</button>
					<br/><br/>
                    </p>
				</form>
               </div>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</body>
</html>
