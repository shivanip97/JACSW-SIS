<% 
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
	        window.location = "https://socialwork.cc.uic.edu/SIS/MSWApplicationReports.asp?ID=" + sel.value;
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
<!--#include file="headerApplicant.asp"-->
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
      <br /><br /><br />
      <%admterm=Request("ID") %>
    <table>
        <tr>
         <td>  <button type="submit" name="Button8" onclick="this.form.action='AdmissionReport.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button8">Report 1 - Overall Summary Status - Count</</button></td>
         <td>  <button type="submit" name="Button4" onclick="this.form.action='MPHReport.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button4">Report 5 - MPH Students</</button> </td>   
         <td>  <button type="submit" name="Button2" onclick="this.form.action='AdmissionReportApplied.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button2">Report 9 - Applied</</button> </td>
         <td>  <button type="submit" name="Button10" onclick="this.form.action='RaceGenderReportApplicant.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button10">Report 13 - Race/Gender Ethnicity Report - Applicant</</button> </td>
        </tr>

        <tr>
         <td>  <button type="submit" name="Button12" onclick="this.form.action='RaceGenderReportConfirmed.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button12">Report 2 - Race/Gender Ethnicity Report - Confirmed</</button> </td>
         <td>  <button type="submit" name="Button6" onclick="this.form.action='InternationalReport.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button6">Report 6 - International Students</</button> </td>
         <td>  <button type="submit" name="Button3" onclick="this.form.action='AdmissionReportWaitlist.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button3">Report 10 - Accept - Waitlist</</button> </td>
         <td>  <button type="submit" name="Button11" onclick="this.form.action='RaceGenderReportAccepted.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button11">Report 14 - Race/Gender Ethnicity Report - Accepted</</button> </td>
        </tr>

        <tr>
         <td>  <button type="submit" name="Button9" onclick="this.form.action='Report.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button9">Report 3 - Admission Report - Confirm</</button></td>
         <td>  <button type="submit" name="Button7" onclick="this.form.action='AcceptReport.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button7">Report 7 - Accept Report - ADV FT PM TR</</button</td> 
         <td>  <button type="submit" name="Button5" onclick="this.form.action='DeferredReport.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button5">Report 11 - Deferred Students</</button> </td>
         <td>  <button type="submit" name="Button13" onclick="this.form.action='AdmissionReportPriorCollegeApplied.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button13">Report 15 - Admission Report - Prior College - Applied</</button> </td>
        </tr>

        <tr>
         <td>  <button type="submit" name="Button14" onclick="this.form.action='AdmissionReportPriorCollegeConfirmed.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button14">Report 4 - Admission Report - Prior College - Confirmed</</button> </td>
         <td>  <button type="submit" name="Button16" onclick="this.form.action='ApplicantDeposit.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button16">Report 8 - Deposit Report</</button> </td>
         <td>  <button type="submit" name="Button1" onclick="this.form.action='DenialReport.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button1">Report 12 - Denial</</button> </td>  
         <td>  <button type="submit" name="Button15" onclick="this.form.action='fiveyearReportPriorCollegeConfirmed.asp?term=<% Response.Write(admterm) %>'; this.forms.submit();" id="Button15">Report 16 - Trend Report - Prior College - Confirmed</</button> </td>
        </tr>
    
</table>

     
      
      
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