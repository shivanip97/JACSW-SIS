<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
user_name = Request("UN")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | Add New Role</title>
<link rel="stylesheet" href="css/newUserStyle.css" type="text/css" media="screen" />
 <script>     function Trim(str) { return str.replace(/^\s*|\s*$/g, ""); } </script>
<script language="javascript">
    function validate() {
        return true;
    }
</script>
</head>
<body>
    <div id="content" align=center>
            <div id="steps">
				<form id="form1" method="post" action="AfterAddRole.asp">
					<h3>Add New Role for User <% Response.Write(user_name) %></h3>
                     <br/>
                    <a href='UserRoles.asp?UN=<% Response.Write(user_name) %>'>Back to Show Roles</a> 
                    <br/> <br/>
                    <p>
                    <label>Add Role Form</label>
                    <br/><br/><br/>
                     <label>User Name</label>
					<input type="text" name="userName" id="userName" value="<% Response.Write(user_name) %>" readonly=true/> 
                    <br/><br/><br/>
                    <label>Program Type</label>
					<select name="programType">
                    <option value="MSW">MSW</option>
                    <option value="PHD">PHD</option>
                     <option value="T73">T73</option>
                    <option value="MPH">MPH</option>
                    </select>
                    <br/><br/><br/>
                    <label>Role Type</label>
					<select name="roleType">
                    <option value="Application">Application</option>
                    <option value="Current">Current</option>
                    <option value="Field">Field</option>
                    <option value="Agency">Agency</option>
                    </select> 
                    <br/><br/><br/>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/><br/><br/>
					<button type="submit" name="Submit" onclick="return validate();">Add Role</button>
					<br/><br/>
                    </p>
				</form>
               </div>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</body>
</html>
