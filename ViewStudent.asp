<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UIN = Request("Button1")

set rs=Server.CreateObject("ADODB.recordset")
query="select * from Applicants where UIN ='"& UIN &"' and Term_CD = '"& Request("ID") &"'"
rs.Open query,conn
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | View Student</title>
<link rel="stylesheet" href="css/tabStyle.css" type="text/css" media="screen" />
<link rel="stylesheet" href="css/screen.css" />
<script type="text/javascript" src="jquery/jquery-1.9.0.js"></script>
<script type="text/javascript" src="jquery/jquery.validate.js"></script>
<script src="jquery/jquery.mask.js" type="text/javascript"></script>
<script type="text/javascript" src="jquery/sliding.form.js"></script>
<script type="text/javascript" src="jquery/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
	$(document).ready(function () {
		$('.date').mask('00/00/0000');
		$('.homephone').mask('(000) 000-0000');
		$('.workphone').mask('(000) 000-0000 x00000');
		$('.iphone').mask('+000 000 000 000');
		$('.gpa').mask('0.00');

		// Reset Checkbox values
		$('.checkboxField').each(function () {
			if ($(this).val() == "Y") {
				$(this).attr('checked', true);
			}
			else {
				$(this).attr('checked', false);
			}


		});

		//$('input.rbfield').removeAttr('checked');
		var checkedElm = $('input.rbfieldhidden').val();
		if (checkedElm != '' || checkedElm != undefined) {
			$('input:radio[value=' + checkedElm + ']').attr('checked', 'checked');
		}

		var fieldval = $('input.cbhidden').val();
		if (fieldval != '' || fieldval != undefined) {
			$('input.cbft[value=' + fieldval + ']').attr('checked', 'checked');
		}


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
		else {
			if ($("#gender option:selected").val() == "0") {
				shouldProceed = false;
				alert('Please select a gender');
			}
		}
		return shouldProceed;
	}
	</script>

</head>
<body>
	<div id="content" >
	<h3>View MSW Student Information</h3>
					 <br/>
					<a href="ShowStudents.asp?ID=220178">Back to Show Students</a> 
					<br/> <br/>
					<% If Session("Username") = "dnmiles" or Session("Username") = "test" or Session("Username") = "tmorri3" or Session("Username") = "apradh6" or Session("Username") = "jrich" Then %>
					<button style="border:none;outline:none;-moz-border-radius: 10px;-webkit-border-radius: 10px;font-weight:bold;margin: 0px auto;clear:both;padding: 7px 25px;font-size:22px;display: block; background:#4797ED;font-family:Century Gothic, Helvetica, sans-serif;" type="submit" name="Button1" onclick="studentForm.action='EditStudent.asp?UIN=' + this.value; studentForm.submit();" value='<% Response.write rs("UIN") %>'>Edit Student</button><br /><br />
					<% End If %>
					
	<div id="wrapper">
	 <div id="navigation" style="display: none;">
					<ul>
						<li><a href="#">Student Demographics</a></li>
						<li><a href="#">Application Info</a></li>
						<li><a href="#">Scholarship Info</a></li>
					</ul>
			  </div>
	<div id="steps" align="center">
				<form id="studentForm" method="post" action="EditStudent.asp">
				
					
				
			  
				<fieldset class="step">
				<legend></legend>

				<p>
					<br/><br/><br/>
					<label>First Name</label>
					<input type="text" name="FirstName" required id="FirstName" value='<%Response.write rs("FirstName") %>' readonly=true/>   
					<label>Middle Name</label>
					<input type="text" name="middlename" id="middlename" value='<%Response.write rs("middlename") %>' readonly=true/> 
					<label>Last Name</label>
					<input type="text" name="lastname" required id="lastname" value='<%Response.write rs("lastname") %>' readonly=true/>    
					<br/><br/><br/>
					<label>Alternate Name</label>
					<input type="text" name="maidenname" id="maidenname" value='<%Response.write rs("maidenname") %>' readonly=true/>
					<label>UIN</label>
					<input type="text" name="uin"  id="uin" value='<%Response.write rs("uin") %>' readonly=true/> 
					<label>Date of Birth</label>
					<input type="text" name="dob" class="date"  id="dob" value='<%Response.write rs("DateOfBirth") %>' readonly=true/> 
					<br/><br/><br/>
					<label>Gender</label>
					<input type="text" name="gender"  id="gender" value='<%Response.write rs("gender") %>' readonly=true/>
					<label>Race/Ethnicity</label>
					<input type="text" name="Race_ethinicity"  id="Race_ethinicity" value='<%Response.write rs("Race_ethinicity") %>' readonly=true/>                                                           
					<label>Race SubCategory</label>
					<input type="text" name="race_subcategory" id="race_subcategory" value='<%Response.write rs("race_subcategory") %>' readonly=true/>
					<br/><br/><br/><br />               
					<label>Email</label>
					<input type="text" name="email" id="email" value='<%Response.write rs("email") %>' readonly=true/>
					<label>Home Phone</label>
					<input type="text" name="homephone" class="homephone"  id="homephone" value='<%Response.write rs("homephone") %>' readonly=true/>
					<label>Cell Phone</label>
					<input type="text" name="workphone" class="workphone"  id="workphone" value='<%Response.write rs("workphone") %>' readonly=true/>                             
					
					<br/><br/><br/>
					<label>Mailing Address 1</label>
					<input type="text" name="currentAddress1"  id="currentAddress1" value='<%Response.write rs("currentAddress1") %>' readonly=true/>
					<label>Mailing Address 2</label>
					<input type="text" name="currentAddress2" id="currentAddress2" value='<%Response.write rs("currentAddress2") %>' readonly=true/>
					<label>Mailing City</label>
					<input type="text" name="currentcity"  id="currentcity" value='<%Response.write rs("currentcity") %>' readonly=true />
					<br/><br/><br/><br/>
					<label>Mailing State</label>
					<input type="text" name="currentstate"  id="currentstate" value='<%Response.write rs("currentstate") %>' readonly=true/>
					<label>Mailing Zip</label>
					<input type="text" name="currentzip" class="zip"  id="currentzip" value='<%Response.write rs("currentzipcode") %>' readonly=true/>
					<label>Country</label>
					<input type="text" name="currentcountry"  id="currentcountry" value='<%Response.write rs("currentcountry") %>' readonly=true/>
					<br/><br/><br/><br/>
					<label>Country of Origin</label>
					<input type="text" name="OriginCountry" id="OriginCountry" value='<%Response.write rs("OriginCountry") %>' readonly="true"/>
                    <label>International</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;"  name="International" id="International" class="checkboxField" value='<%Response.write rs("International") %>' />
					<br/><br/><br/><br/>
					
					</p>
					</fieldset>

					<div id="application" align="center">
					<fieldset class="step">
					<legend></legend>
				
					<p>
					<br/>
					<label>Application Status</label>
					<input type="text" name="application_status" id="application_status" value='<%Response.write rs("application_status") %>' readonly="true"/>
					<label >Admission Decision</label>
					<input type="text" name="admission_decision" id="admission_decision" value='<%Response.write rs("admission_decision") %>' readonly="true"/>
				   
					<label>Ready for Review Date</label>
					<input type="text" name="readyforreviewdate" id="readyforreviewdate" class="date" value='<%Response.write rs("readyforreviewdate") %>' readonly="true"/>                      
					
					
					<br/><br/><br/><br/>
					<label>OAR Application Date</label>
					<input type="text" name="oar_application_date" class="date" required id="oar_application_date" value='<%Response.write rs("oar_application_date") %>' readonly=true/>
					<label>Degree Program</label>
					<input type="text" name="Degree_Program" id="Degree_Program" value='<%Response.write rs("Degree_Program") %>' readonly="true"/> &nbsp;&nbsp;&nbsp;&nbsp
					<label>Field Type</label>
					<input type='hidden' style="margin:0;width:20px;height:20px;" class="cbhidden" value='<%Response.write rs("Field_Type") %>' />
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="Field_Type" class="cbft" value='F' />
					<label class="clearWidth">Generalist</label>
					<input type="checkbox"  disabled="disabled" style="margin:0;width:20px;height:20px;" name="Field_Type"  class="cbft" value='C' />
					<label class="clearWidth">Specialization</label>
					<br /><br /><br /><br />
					<label>Program Type</label>
					<input type="text" name="program_type" id="program_type" value='<%Response.write rs("program_type") %>' readonly="true"/>
					
					<label style="padding-left:40px">Admitted to School</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="admitted_to_school" id="admitted_to_school" class="checkboxField" value='<%Response.write rs("admitted_to_school") %>' />

					
					<label>Specialization</label>
					<input type="text" name="concentration" id="concentration" value='<%Response.write rs("concentration") %>' readonly="true"/>
					  
					<br /><br /><br /><br />
					
					<label>Decision Date</label>
					<input type="text" name="decision_dt" id="decision_dt" class="date" value='<%Response.write rs("decision_dt") %>' />
					<label>Decision Letter Sent Date</label>
					<input type="text" name="Decision_Letter_Sent_Date" id="Decision_Letter_Sent_Date" class="date" value='<%Response.write rs("Decision_Letter_Sent_Date") %>'/>
					<label>Confirmed Due Date</label>
					<input type="text" name="ConfirmedDueDate" id="ConfirmedDueDate" class="date" value='<%Response.write rs("ConfirmedDueDate") %>' />
					
					<br /><br /><br /><br /><br />                   
					 
					<label>Confirmed Date</label>
					<input type="text" name="Confirmed_Dt" id="Confirmed_Dt" class="date" value='<%Response.write rs("Confirmed_Dt") %>' />
					<label>Confirmed</label>
					<input type="text" name="Confirmed" id="Confirmed" value='<%Response.write rs("Confirmed") %>' readonly="true"/>                   
					<label>Admit Term</label>
					<input type="text" name="Admit_Term" id="Admit_Term" value='<%Response.write rs("Admit_Term") %>' readonly="true"/>
					<br /><br /><br /><br />
					 <label>Withdraw Status</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;margin-right: 180px;" name="withdrawn" id="withdrawn" class="checkboxField" value='<%Response.write rs("withdrawn") %>' />
					<label>Withdraw Reason</label>
					<input type="text" name="withdraw_reason" id="withdraw_reason" value='<%Response.write rs("withdraw_reason") %>' readonly="true"/>
					<label>Withdrawn Date</label>
					<input type="text" name="WithdrawnDate" id="WithdrawnDate"  class="date" value='<%Response.write rs("WithdrawnDate") %>' readonly=true/>
					
					<br /><br /><br /><br />
					<label>Reapplicant</label>
					<input type="text" name="reapplicant" id="reapplicant" value='<%Response.write rs("reapplicant") %>' readonly="true"/>
				   
					<label>Entered By</label>
					<input type="text" name="enteredby"  id="enteredby" value='<%Response.write rs("enteredby") %>' readonly="true"/>
					<label>Last Updated Date</label>
					<input type="text" name="LastUpdatedDt" id="LastUpdatedDt"  value='<%Response.write rs("LastUpdatedDt") %>' readonly=true/>

					<br /><br /><br /><br />
					<label>Limited Status</label>
					<input type="text" name="Limited_status" id="Limited_status" value='<%Response.write rs("Limited_status") %>' readonly="true" /> 
					<label>Received Deposit</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="received_deposit" id="received_deposit" class="checkboxField" value='<%Response.write rs("received_deposit") %>' />

					<label>Forward to Field</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="forward_to_field" id="forward_to_field" class="checkboxField" value='<%Response.write rs("forward_to_field") %>' />
					
					
					<br /><br /><br /><br />
					<label>Deferred From</label>
					<input type="text" name="DeferredFrom" id="DeferredFrom"  value='<%Response.write rs("DeferredFrom") %>' readonly=true/>
					<label>Deferred To</label>
					<input type="text" name="DeferredTo" id="DeferredTo"  value='<%Response.write rs("DeferredTo") %>' readonly=true/>
                    <label>Adv Verification</label>
					<input type="checkbox" disabled="disabled" name="Adv_verification" style="margin:0;width:20px;height:20px;margin-right: 140px;" class="checkboxField" id="Adv_Verification"  value='<%Response.write rs("Adv_Verification") %>' />
					
					
					<br /><br /><br /><br />
					<label>Comments</label>
					<textarea id="comments" name="comments" cols="70" rows="5" readonly="true"><%Response.write rs("Comments") %></textarea> 
					</p>
					<br /><br /><br /><br />
					<p>
					
					<label>Credit In Statistics</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="Credit_in_Statistics" id="Credit_in_Statistics" class="checkboxField" value='<%Response.write rs("Credit_in_Statistics") %>'/>
					<label>Credit in BA BS</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="credit_in_ba_bs" id="credit_in_ba_bs" class="checkboxField" value='<%Response.write rs("credit_in_ba_bs") %>'/>
					<label>Credit in English</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="credit_in_english" id="credit_in_english" class="checkboxField" value='<%Response.write rs("credit_in_english") %>' />
					<br /><br /><br /><br />                    
					
					<label>Requesting Schools</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="requesting_schools" id="requesting_schools" class="checkboxField" value='<%Response.write rs("requesting_schools") %>' />
					<label>Financial Aid Request</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="Financial_Aid_Request" id="Financial_Aid_Request" class="checkboxField" value='<%Response.write rs("Financial_Aid_Request") %>'/>
					<br/><br/><br/><br/>
					<label>Basic Skills TAP</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;"  name="Basic_Skill_Test" id="Basic_Skill_Test" class="checkboxField" value='<%Response.write rs("Basic_Skill_Test") %>'/>
					<label>Basic Skills ACT/SAT</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;"  name="ACT_SAT" id="ACT_SAT" class="checkboxField" value='<%Response.write rs("ACT_SAT") %>' />
					<label>Passed Test</label>
					<input type="text" name="Passed_Test" id="Passed_Test" value='<%Response.write rs("Passed_Test") %>' readonly=true/>
					<label>UG College</label>
					<input type="text" name="ugcollege" id="ugcollege" value='<%Response.write rs("ugcollege") %>' readonly=true/>

					<br/><br/><br/><br/>
					<label>UG GPA</label>
					<input type="text" name="uggpa" class="gpa" id="uggpa" value='<%Response.write rs("UGGPA") %>' readonly=true/>
					<label>UG Major</label>
					<input type="text" name="ugmajor" id="ugmajor" value='<%Response.write rs("UGMajor") %>' readonly=true/>                    
					<label>Grad College</label>
					<input type="text" name="gradcollege" id="gradcollege" value='<%Response.write rs("gradcollege") %>' readonly=true/>
				
					<br/><br/><br/><br/>
					<label>Grad GPA</label>
					<input type="text" name="gradgpa" class="gpa" id="gradgpa" value='<%Response.write rs("GradGPA") %>' readonly=true/>
					<label>Grad Major</label>
					<input type="text" name="gradmajor" id="gradmajor" value='<%Response.write rs("GradMajor") %>' readonly=true/>
					<label>Grad Degree</label>
					<input type="text" name="graddegree" id="graddegree" value='<%Response.write rs("GradDegree") %>' readonly=true/>
					
					<br/><br/><br/><br/>
					</p>
					</fieldset>
					
					<div id="scholarship" align="center">
					<fieldset class="step">
					<legend></legend>
					<p>
					<br/>

					<label>Scholarship Type 1</label>
					<input type="text" name="award_type1" id="award_type1" value='<%Response.write rs("award_type") %>' readonly=true/>
					   
					<label>Award Amount</label>
					<input type="text" name="award_amount1" id="award_amount1" value='<%Response.write rs("award_amount") %>' readonly=true/>
					
					<label>Award Date</label>
					<input type="text" name="award_date1" id="award_date1" class="date" value='<%Response.write rs("award_date") %>' readonly=true/>
					<br/><br/><br/><br/>

					<label>Scholarship Type 2</label>
					<input type="text" name="award_type2" id="award_type2" value='<%Response.write rs("award_type2") %>' readonly=true/>
					   
					<label>Award Amount</label>
					<input type="text" name="award_amount2" id="award_amount2" value='<%Response.write rs("award_amount2") %>' readonly=true/>
					
					<label>Award Date</label>
					<input type="text" name="award_date2" id="award_date2" class="date" value='<%Response.write rs("award_date2") %>' readonly=true/>
					<br/><br/><br/><br/>

					<label>Scholarship Type 3</label>
					<input type="text" name="award_type3" id="award_type3" value='<%Response.write rs("award_type3") %>' readonly=true/>
					   
					<label>Award Amount</label>
					<input type="text" name="award_amount3" id="award_amount3" value='<%Response.write rs("award_amount3") %>' readonly=true/>
					
					<label>Award Date</label>
					<input type="text" name="award_date3" id="award_date3" class="date" value='<%Response.write rs("award_date3") %>' readonly=true/>
					<br/><br/><br/><br/>
					</p>
				  
			
					

					<strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
					<br/>
					 
					</fieldset>
				   

				</form>
				</div>
			   </div>
			   
			</div>
			<br/>
			<!--#include file="footer.asp"-->
</body>
</html>
