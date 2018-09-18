<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UIN = Request("UIN")
Session("UIN") = UIN
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | Add New Record</title>
<link rel="stylesheet" href="css/tabStyle.css" type="text/css" media="screen" />
<link rel="stylesheet" href="css/screen.css" />
<script type="text/javascript" src="jquery/jquery-1.9.0.js"></script>
<script type="text/javascript" src="jquery/jquery.validate.js"></script>
<script src="jquery/jquery.mask.js" type="text/javascript"></script>
<script type="text/javascript">
    $(document).ready(function () {
        $('.date').mask('00/00/0000');
    });

    function validate() {
        var shouldProceed = true;
        $('#studentForm').find(':input:not(button)').each(function () {
            var $this = $(this);
            var valueLength = jQuery.trim($this.val()).length;
            if ($(this).attr("required") && $(this).val() === "") {
                shouldProceed = false;
                $this.css('background-color', '#FFEDEF');
            }
            else
                $this.css('background-color', '#FFFFFF');
        });
        if (shouldProceed == false) {
            alert('Please Complete form by filling in fields highlighted in Red.')
        }
        return shouldProceed;
    }
 	</script>

</head>
<body>
    <div id="content" align=center>
            <div id="steps">
				<form id="studentForm" method="post" action="AfterAddRecord.asp">
					<h3>Add New Record</h3>
                     <br/>
                    <a href=ShowCurrentStudents.asp>Back to Show Current Students</a> 
                    <br/> <br/>
                    <p>
                    <label>Add Record Form</label>
                    <br/><br/><br/>
                    <label>Degree Program</label>
   	                <select name="DegreeProgram" id="DegreeProgram">
         			<option value="MSW">MSW</option>
  					<option value="PHD">PHD</option>
				    </select>
                    <label>Limited Status</label>
   	                <select name="LimitedStatus" id="LimitedStatus">
         			<option value="Yes">Yes</option>
  					<option value="No">No</option>
				    </select>
                    <label>Program Type</label>
   	                <select name="ProgramType" id="ProgramType">
                    <option value="0">-- Select --</option>
         			<option value="FT">FT</option>
  					<option value="PM">PM</option>
					<option value="Adv">Adv</option>
                    <option value="TR">TR</option>   
                    <option value="TR-FT">TR-FT</option>
                    <option value="TR-PM">TR-PM</option>
                    <option value="MPH">MPH</option>
                    <option value="MPH-FT">MPH-FT</option>
                    <option value="MPH-PM">MPH-PM</option>
				    </select>
                    <br/><br/><br/>
                    <label>Concentration</label>
   	                <select name="Concentration" id="Concentration">
         			<option value="0">-- Select --</option>
  					<option value="CAP">CAP</option>
					<option value="CHF">CHF</option>
                    <option value="CHUD">CHUD</option>
					<option value="SCH">SCH</option>
                    <option value="MH">MH</option>
					<option value="HLT">HLT</option>
				    </select>
                     <label>Decision</label>
   	                <select name="Decision" id="Decision">
         			<option value="0">-- Select --</option>
  					<option value="A">A</option>
					<option value="D">D</option>
                    <option value="S">S</option>
					<option value="N">N</option>
                    <option value="AR">AR</option>
					<option value="DF">DF</option>
                    <option value="W">W</option>
					<option value="WD">WD</option>
                    <option value="AWD">AWD</option>
					<option value="SWD">SWD</option>
                    <option value="NWD">NWD</option>
					<option value="ND">ND</option>
                    <option value="ARWD">ARWD</option>
					<option value="WWD">WWD</option>
                    <option value="IN">IN</option>
				    </select>
                    <label>Confirmed</label>
   	                <select name="Confirmed" id="Confirmed">
         			<option value="Yes">Yes</option>
  					<option value="No">No</option>
				    </select>
                    <br/><br/><br/>
                    <label>Confirmed Date</label>
					<input type="text" name="date" class="date" required id="date"/> 
                    <label>Admit Term</label>
					<input type="text" name="AdmitTerm" required id="AdmitTerm" />
                    <label>Track</label>
					<input type="text" name="Track" required id="Track" />
                    <br/><br/><br/>
                    <label>Current Year</label>
					<input type="text" name="currentyear" required id="currentyear" />
                    <label>Advisor</label>
					<input type="text" name="Advisor" required id="Advisor" />
                    <label>Applying For Graduation?</label>
   	                <select name="ApplyingForGraduation" id="ApplyingForGraduation">
         			<option value="0">Yes</option>
  					<option value="1">No</option>
				    </select>
                    <br/><br/><br/><br/>
                     <label>Graduation Term Applied For</label>
					<input type="text" name="GraduationTermAppliedFor" required id="GraduationTermAppliedFor" />
                    <label>Term Graduated</label>
					<input type="text" name="TermGraduated" required id="TermGraduated" />
                    <label>Degree Applying For</label>
					<input type="text" name="DegreeApplyingFor" required id="DegreeApplyingFor" />
                    <br/><br/><br/><br/>
                    <label>Mailbox Number</label>
					<input type="text" name="MailboxNumber" id="MailboxNumber" />
                    <br/><br/><br/>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/><br/><br/>
					<button type="submit" name="Submit" onclick="return validate();">Add Record</button>
					<br/><br/>
                    </p>
				</form>
               </div>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</body>
</html>
