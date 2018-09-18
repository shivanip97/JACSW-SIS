﻿<% 
ErrMsg = Request("ErrMsg")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>SIS | STUDENTS</title>
<link rel="stylesheet" href="css/tabstyle.css" type="text/css" />
    <script type="text/javascript" src="jquery/jquery-1.9.0.js"></script>
    <script type="text/javascript" src="jquery/filtertable.js"></script>
    <script type="text/javascript" src="jquery/jquery.jeditable.mini.js"></script>
	<script type="text/javascript">
	    $(document).ready(function () {
	        $('.edit').editable('UpdateStudent.asp');
	        $('.editableGender').editable('UpdateStudent.asp', {
	            data: " {'M':'M','F':'F', 'selected':'M'}",
	            type: 'select',
	            submit: 'OK'
	        });

	    });

	    function getval(sel) {
	        window.location = "https://socialwork.cc.uic.edu/SIS/PhDApplicationReports.asp?ID=" + sel.value;
	    }

	    
 	</script>
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
	</style>
 </head>
<body bgcolor="#f2f2f2">
<!--#include file="header.asp"-->
<!--#include file="headerPHDApplicant.asp"-->
<br /><br />
<div align="center"><form action="" method="post"> 
     <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
     <br/> <br/>

     <label>Admit Term</label>
      <select name="admit_term" id="admit_term" onchange="getval(this);">
     <%
                        query = "select Term_CD,Admit_Term from AdmitTerm_Codes"
                        set drs = conn.execute(query)
                        do while not drs.eof 
                       if Request("ID") = drs.Fields(0) then
                        %>
                        <option value="<%= drs.Fields(0) %>" selected="selected"><%= drs.Fields(1) %></option>
                        <% 
                        else
                        %>
                        <option value="<%= drs.Fields(0) %>"><%= drs.Fields(1) %></option>
                        <% end if            
                        drs.MoveNext
                        Loop
                    
                    %>
                     </select>
      <br /><br />
      <%admterm=Request("ID") %>
      <button type="submit" name="Button1" onclick="this.form.action='PHDReport.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button1">Admission Report - Confirms</</button> 
    <button type="submit" name="Button2" onclick="this.form.action='PHDAdmissionReport.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button2">Admission Report - Count</</button> 
    <button type="submit" name="Button3" onclick="this.form.action='PHDRaceGenderReportApplicant.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button3">Race/Gender Ethinicity Report - Applicant</</button> 
    <button type="submit" name="Button4" onclick="this.form.action='PHDRaceGenderReportAccepted.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button4">Race/Gender Ethinicity Report - Accepted</</button> 
    <button type="submit" name="Button5" onclick="this.form.action='PHDRaceGenderReportConfirmed.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button5">Race/Gender Ethinicity Report - Confirmed</</button> 
    
    <button type="submit" name="Button7" onclick="this.form.action='PHDDenialReport.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button7">Denial</</button> 
    <button type="submit" name="Button8" onclick="this.form.action='PHDAdmissionReportWaitlist.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button8">Accept - Waitlist</</button> 
    <button type="submit" name="Button9" onclick="this.form.action='PHDAdmissionReportApplied.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button9">Applied</</button> 


     
      
      
      <br /><br />
    

</form> 

</div>
<!-- overlayed element -->
<div class="apple_overlay" id="overlay">
  <!-- the external content is loaded inside this tag -->
  <div class="contentWrap"></div>
</div>
<p>&nbsp;</p>

</body>
<!--#include file="footer.asp"-->
</html>