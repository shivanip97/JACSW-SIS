<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
Dim msg
submitButton=Request("Submit")="Submit"

if submitButton then

	Session("AccessLevel") = "2"
	update_form_query = "update Users set status=?, AccessLevel=? where Username=? and Year=?"
	
	Set objCommand = Server.CreateObject("ADODB.Command") 
	
	objCommand.ActiveConnection = conn
	objCommand.CommandText = update_form_query


	
	objCommand.Parameters(0).value = "Submitted"
	objCommand.Parameters(1).value = "2"
	objCommand.Parameters(2).value = Session("Username")
	objCommand.Parameters(3).value = Application("currentYear")

	Set objRS = objCommand.Execute()
	
	update_form_query = "update FAAR set Date=? where Username=? and Year=?"
	
	Set objCommand = Server.CreateObject("ADODB.Command") 
	
	objCommand.ActiveConnection = conn
	objCommand.CommandText = update_form_query
	sub_info = Date() & " - " & Time()

	
	objCommand.Parameters(0).value = sub_info
	objCommand.Parameters(1).value = Session("Username")
	objCommand.Parameters(2).value = Application("currentYear")

	Set objRS = objCommand.Execute()
	
	Response.Redirect "Submitted.asp"

End If	



msg = "Hi" & " " & Session("Username")
set rs=Server.CreateObject("ADODB.recordset")
query="select * from Users where Username='"& Session("Username") &"' and Year="&Application("currentYear")
rs.Open query,conn
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<link rel="stylesheet" href="images/style.css" type="text/css" />
	<title>FAAR | UIC</title>	
<script type="text/JavaScript">
	<!--
	function timedRefresh(timeoutPeriod) {
		setTimeout("location.reload(true);",timeoutPeriod);
	}
	//   -->
</script>
</head>
<body >
	<div id="page" align="center">
		<div id="header">
		   <!--#include file="header.asp"-->
			<div align="right" class="links_menu" id="menu">
           <!--#include file="Controls.asp"-->
             </div>
		</div>
		<br /><br />

		<div id="content">
			<div id="leftpanel">
				<div class="table_top">
					<div align="center">
						<span class="title_panel">Updates</span>
					</div>
				</div>
				<div class="table_content">
					<div class="table_text">
						
						<span class="news_date"><% Response.Write(msg) %> </span>
                        <br /><br />
                        
						<span class="news_text"> Your Last Login was  </span><br />
						<span class="news_date"><% Response.write rs("LastLogin")%></span> <br />
							<% rs.close 
                            conn.close%>								
					</div>
				</div>
				<div class="table_bottom">
					<img src="images/table_bottom.jpg" width="204" height="23" border="0" alt="" />
				</div>
				<br/>			
			</div>
			<form id="form1" method="post" action="">
			<div id="stylized" class="myform" align=center >
			 <h1><a href="pdf_gen.asp" style="font-size:12pt; color:orange; " target="_blank"> Click Here</a> to view/print your Faculty Annual Activity Report before submission.</h1>
			
		    <div class="spacer"></div> <br/>
    <button type="submit" name="Submit" value="Submit" >Submit</button>
    <div class="spacer"></div><br/>
<span class="PDFtext">(*Note: Forms once submitted cannot be edited again)</span>
			</div></form>
			<div class="spacer"></div><br />
			<div class="spacer"></div>
		  <div class="footer"><br />
            <!--#include file="footer.asp"-->
		  </div>
          <p class="news_date">&nbsp;</p>
          <p class="copyright">Powered by Web Services, UIC</p>
      </div>
	</div>
</body>
</html>


