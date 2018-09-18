<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<title>SIS | Change Password</title>
<link rel="stylesheet" href="css/loginStyle.css" type="text/css" media="screen" />
<script>    function Trim(str) { return str.replace(/^\s*|\s*$/g, ""); } </script>
<script language="javascript">
    var request = makeObject();
    function validate() {
        var password = document.getElementById("password");
        if (Trim(password.value) == '') {
            alert("Please provide a password");
            password.focus();
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
				<form id="form1" method="post" action="ChangePass.asp">
					<h2>Enter new Password</h2>
                    <br/>
                    <p>   
                    <br/><br/><br/>
					<label>Password </label>
					<input type="password" name="password" id="password" />
                    <br/><br/><br/>
					<button type="submit" name="Submit" onclick="return validate();">Change Password</button>
					<br/>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong><br />
                    </p>
				</form>
                <br />

          
               </div>
			</div>
            	<div class="footer"><br />
            <!--#include file="footer.asp"-->
		  </div>
            </div>
</body>
</html>