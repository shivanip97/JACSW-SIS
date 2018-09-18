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
	        window.location = "https://socialwork.cc.uic.edu/SIS/MSWCurrentReports.asp?ID=" + sel.value;
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
<!--#include file="headerCurrentStudent.asp"-->
<br /><br />
<div align="center"><form action="" method="post"> 
     <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
     <br/> <br/>
    <table>
        <tr>
           <td> <button type="submit" name="Button1" onclick="this.form.action='CurrentStudentReport.asp'; this.forms.submit();" id="Button1">Current Students</</button> </td>
           <td> <button type="submit" name="Button10" onclick="this.form.action='CurrentADVAcceptReport.asp'; this.forms.submit();" id="Button10">Accept Report - ADV </</button></td>
           <td> <button type="submit" name="Button11" onclick="this.form.action='CurrentConcentrationPTReport.asp'; this.forms.submit();" id="Button11">Concentration Change Report </</button> </td>
        </tr>

        <tr>
           <td> <button type="submit" name="Button2" onclick="this.form.action='CurrentMPHReport.asp'; this.forms.submit();" id="Button2">MPH Students</</button> </td>
           <td> <button type="submit" name="Button9" onclick="this.form.action='OffTrackReport.asp'; this.forms.submit();" id="Button9">Off Track Report</</button></td>
           <td> <button type="submit" name="Button12" onclick="this.form.action='Current10dayWithdrawalPrior.asp'; this.forms.submit();" id="Button12">Withdrawal Report after 10 days </</button> </td>
        </tr>

        <tr>
           <td> <button type="submit" name="Button3" onclick="this.form.action='CurrentStudentOneAdvisorReport.asp'; this.forms.submit();" id="Button3">One Advisor Report</</button></td>
           <td> <button type="submit" name="Button8" onclick="this.form.action='LOAReport.asp'; this.forms.submit();" id="Button8">LOA Report</</button></td>
           <td> <button type="submit" name="Button13" onclick="this.form.action='CSWE Report.asp'; this.forms.submit();" id="Button13">CSWE Report </</button>  </td>
        </tr>

        <tr>
           <td> <button type="submit" name="Button4" onclick="this.form.action='CurrentStudentAllAdvisorReport.asp'; this.forms.submit();" id="Button4">All Advisor Report</</button></td>
           <td> <button type="submit" name="Button7" onclick="this.form.action='GradTermAppliedFor.asp'; this.forms.submit();" id="Button7">Graduation Applied Report</</button></td>
           <td> <button type="submit" name="Button14" onclick="this.form.action='DegreeCompletionDurationReport.asp'; this.forms.submit();" id="Button14">Degree Completion Duration Report </</button> </td>
        </tr>

        <tr>
           <td> <button type="submit" name="Button5" onclick="this.form.action='rptAcademicAffairsReport.asp'; this.forms.submit();" id="Button5">Academic Affairs Report</</button> </td>
           <td> <button type="submit" name="Button6" onclick="this.form.action='GraduationReport.asp'; this.forms.submit();" id="Button6">Graduation Report</</button></td>
           <td></td>
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