  <!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UIN = Request("UIN")
Admtrm=Request("Admit_Term")
set rs=Server.CreateObject("ADODB.recordset")
query="select * from Applicants where UIN ='"& UIN &"' and Admit_Term='"&Admtrm&"'"
rs.Open query,conn
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | Edit Student</title>
<link rel="stylesheet" href="css/tabStyle.css" type="text/css" media="screen" />
<link rel="stylesheet" href="css/screen.css" />
<script type="text/javascript" src="jquery/jquery-1.9.0.js"></script>
<script type="text/javascript" src="jquery/jquery.validate.js"></script>
<script src="jquery/jquery.mask.js" type="text/javascript"></script>
<script type="text/javascript" src="jquery/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="jquery/sliding.form.js"></script>

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

		$('.checkboxField').on('click', function () {
			if ($(this).is(":checked")) {
				$(this).attr('value', 'Y');
			} else {
				$(this).attr('value', 'N');
			}
		});

		$('input.rbfield').removeAttr('checked');
		var checkedElm = $('input.rbfieldhidden').val();
		if (checkedElm != '' || checkedElm != undefined) {
			$('input:radio[value=' + checkedElm + ']').attr('checked', 'checked');
		}

		var fieldval = $('input.cbhidden').val();
		if (fieldval != '' || fieldval != undefined) {
			$('input.cbft[value=' + fieldval + ']').attr('checked', 'checked');
		}

		$('#withdrawn').change(function () {
			$("#withdraw_reason").prop("disabled", !$(this).is(':checked'));
		});

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
		<a style="font-size:12pt;" href="ShowStudents.asp?ID=220168">Back to Show Students</a>
		<br /><br />
		<br />
		<strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
		<br />
		<div id="wrapper">
		<div id="navigation" style="display: none;">
				<ul>
					<li><a href="#">Student Demographics</a></li>
					<li><a href="#">Application Info</a></li>
					 <li><a href="#">ScholarShip Info</a></li>
				</ul>
			</div>
		
			<div id="steps" align="center">
				<form id="studentForm" method="post" action="AfterEditStudent.asp">
				<fieldset class="step">
				<legend></legend>

					<p>
					<br/><br/><br/>
					<label>First Name</label>
					<input type="text" name="FirstName" required id="FirstName" value="<%Response.write rs("FirstName") %>" readonly=true/>   
					<label>Middle Name</label>
					<input type="text" name="middlename" id="middlename" value="<%Response.write rs("middlename") %>" readonly=true/> 
					<label>Last Name</label>
					<input type="text" name="lastname" required id="lastname" value="<%Response.write rs("lastname") %>" readonly=true/>    
					<br/><br/><br/>
					<label>Alternate Name</label>
					<input type="text" name="maidenname" id="maidenname" value="<%Response.write rs("maidenname") %>"/>
					<label>UIN</label>
					<input type="text" name="uin" required id="uin" value='<%Response.write rs("uin") %>' readonly=true/> 
					<label>Date of Birth</label>
					<input type="text" name="dob" class="date" required id="dob" value='<%Response.write rs("DateOfBirth") %>' readonly=true/> 
					<br/><br/><br/>
					<label>Gender</label>
					<select name="gender" id="gender">
					<option value="<%= rs.Fields(11) %>"><%= rs.Fields(11) %></option>
					<option value="M">M</option>
					<option value="F">F</option>
					</select>
					<label>Race/Ethnicity</label>
					<input type="text" name="Race_ethinicity" required id="Race_ethinicity" value='<%Response.write rs("Race_ethinicity") %>' readonly=true/>                                                           
					<label>Race SubCategory</label>
					<input type="text" name="race_subcategory" id="race_subcategory" value='<%Response.write rs("race_subcategory") %>'/>
					<br/><br/><br/><br />               
					<label>Email</label>
					<input type="text" name="email" id="email" value='<%Response.write rs("email") %>'/>
					<label>Home Phone</label>
					<input type="text" name="homephone" class="homephone"  id="homephone" value='<%Response.write rs("homephone") %>'/>
					<label>Cell Phone</label>
					<input type="text" name="workphone" class="workphone"  id="workphone" value='<%Response.write rs("workphone") %>' />                             
					
					<br/><br/><br/>
					<label>Mailing Address 1</label>
					<input type="text" name="currentAddress1" required id="currentAddress1" value="<%Response.write rs("currentAddress1") %>" />
					<label>Mailing Address 2</label>
					<input type="text" name="currentAddress2" id="currentAddress2" value="<%Response.write rs("currentAddress2") %>" />
					<label>Mailing City</label>
					<input type="text" name="currentcity" required id="currentcity" value='<%Response.write rs("currentcity") %>' />
					<br/><br/><br/><br/>
					<label>Mailing State</label>
					<select name="currentstate" id="currentstate">
					<option value="<%= rs.Fields(24) %>"><%= rs.Fields(24) %></option>    
	<option value="AL">AL</option>
	<option value="AK">AK</option>
	<option value="AZ">AZ</option>
	<option value="AR">AR</option>
	<option value="CA">CA</option>
	<option value="CO">CO</option>
	<option value="CT">CT</option>
	<option value="DE">DE</option>
	<option value="DC">DC</option>
	<option value="FL">FL</option>
	<option value="GA">GA</option>
	<option value="HI">HI</option>
	<option value="ID">ID</option>
	<option value="IL">IL</option>
	<option value="IN">IN</option>
	<option value="IA">IA</option>
	<option value="KS">KS</option>
	<option value="KY">KY</option>
	<option value="LA">LA</option>
	<option value="ME">ME</option>
	<option value="MD">MD</option>
	<option value="MA">MA</option>
	<option value="MI">MI</option>
	<option value="MN">MN</option>
	<option value="MS">MS</option>
	<option value="MO">MO</option>
	<option value="MT">MT</option>
	<option value="NE">NE</option>
	<option value="NV">NV</option>
	<option value="NH">NH</option>
	<option value="NJ">NJ</option>
	<option value="NM">NM</option>
	<option value="NY">NY</option>
	<option value="NC">NC</option>
	<option value="ND">ND</option>
	<option value="OH">OH</option>
	<option value="OK">OK</option>
	<option value="OR">OR</option>
	<option value="PA">PA</option>
	<option value="RI">RI</option>
	<option value="SC">SC</option>
	<option value="SD">SD</option>
	<option value="TN">TN</option>
	<option value="TX">TX</option>
	<option value="UT">UT</option>
	<option value="VT">VT</option>
	<option value="VA">VA</option>
	<option value="WA">WA</option>
	<option value="WV">WV</option>
	<option value="WI">WI</option>
	<option value="WY">WY</option>
	<option value="INT">INT</option>
					</select>
					<label>Mailing Zip</label>
					<input type="text" name="currentzip" class="zip" required id="currentzip" value='<%Response.write rs("currentzipcode") %>'/>
					<label>Country</label>
					<input type="text" name="currentcountry" required id="currentcountry" value='<%Response.write rs("currentcountry") %>'/>
					<br/><br/><br/><br/>
					<label>Country of Origin</label>
					<input type="text" name="OriginCountry" id="OriginCountry" value='<%Response.write rs("OriginCountry") %>' />
                    <label>International</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;"  name="International" id="International" class="checkboxField" value='<%Response.write rs("International") %>' />
					<br/><br/><br/><br/>
					
					</p>
					<button type="submit" name="Submit" onclick="this.form.action='AfterEditStudent.asp?UIN=' + this.value; this.forms.submit();" value='<% Response.write rs("UIN") %>'>Save</button><br /><br />
					</fieldset>



					<div id="application" align="center">
					<fieldset class="step">
					<legend></legend>
				
					<p>
					<br/>
					<label>Application Status</label>
					<select name="application_status" id="application_status">
					<option value="<%=rs.Fields(3) %>"><%= rs.Fields(3) %></option>
					<option value="PD-Pending">PD-Pending</option>
					<option value="IN-Incomplete">IN-Incomplete</option>
					<option value="R-In Review">R-In Review</option>
					<option value="C-Complete">C-Complete</option>
					<option value="ND-No Documents">ND-No Documents</option>
					</select>

					<label >Admission Decision</label>
					<select name="admission_decision" id="admission_decision">
					<option value='<%Response.write rs("admission_decision") %>'><%= rs.Fields(39) %></option>
					<option value=" "> </option>
					<option value="A">A</option>
					<option value="D">D</option>
					<option value="S">S</option>
					<option value="DF">DF</option>
					<option value="W">W</option>
					<option value="SWD">SWD</option>
					<option value="NWD">NWD</option>
					<option value="ReAdmit">ReAdmit</option>
					<option value="S-NWD">S-NWD</option>
					<option value="S-NDWD">S-NDWD</option>
					<option value="S-TRWD">S-TRWD</option>
					<option value="S-TRWWD">S-TRWWD</option>
					</select>

					<label>Ready for Review Date</label>
					<input type="text" name="readyforreviewdate" id="readyforreviewdate" class="date" value='<%Response.write rs("readyforreviewdate") %>'/>                      
					
					
					<br/><br/><br/><br/>
					<label>OAR Application Date</label>
					<input type="text" name="oar_application_date" class="date" required id="oar_application_date" value='<%Response.write rs("oar_application_date") %>' readonly=true/>
					<label>Degree Program</label>
					<input type="text" name="Degree_Program" id="Degree_Program" value='<%Response.write rs("Degree_Program") %>' /> &nbsp;&nbsp;&nbsp;&nbsp
					<label>Field Type</label>
					<input type='hidden' style="margin:0;width:20px;height:20px;" class="cbhidden" value='<%Response.write rs("Field_Type") %>' />
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="Field_Type" class="cbft" value='F'/>
					<label class="clearWidth">Generalist</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="Field_Type"  class="cbft" value='C'/>
					<label class="clearWidth">Specialization</label>
					<br /><br /><br /><br />
					<label>Program Type</label>
					<select name="program_type" id="program_type">
					<option value="<%= rs.Fields(37) %>"><%= rs.Fields(37) %></option>
					<option value="FT">FT</option>
					<option value="PM">PM</option>
					<option value="Adv">Adv</option>
					<option value="MPH-Adv">MPH-Adv</option>
					<option value="MPH-FT">MPH-FT</option>
					<option value="MPH-PM">MPH-PM</option>
					<option value="TR-FT">TR-FT</option>
					<option value="TR-PM">TR-PM</option>
					</select>
					
					<label style="padding-left:40px">Admitted to School</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="admitted_to_school" id="admitted_to_school" class="checkboxField" value='<%Response.write rs("admitted_to_school") %>' />

					
					<label>Specialization</label>
					<select name ="concentration" id="concentration">
					<option value="<%= rs.Fields(38) %>"><%= rs.Fields(38) %></option>
					<option value=""></option>
					<option value="CHF">CHF</option>
					<option value="CHUD">CHUD</option>
					<option value="MH">MH</option>
					<option value="SCH">SCH</option>
                    <option value="OCP">OCP</option>
					</select>
					<br /><br /><br /><br />
					
					<label>Decision Date</label>
					<input type="text" name="decision_dt" id="decision_dt" class="date" value='<%Response.write rs("decision_dt") %>' />
					<label style="width:180px">Decision Letter Sent Date</label>
					<input type="text" name="Decision_Letter_Sent_Date" id="Decision_Letter_Sent_Date" class="date" value='<%Response.write rs("Decision_Letter_Sent_Date") %>'/>
					<label>Confirmed Due Date</label>
					<input type="text" name="ConfirmedDueDate" id="ConfirmedDueDate" class="date" value='<%Response.write rs("ConfirmedDueDate") %>' />
					<br /><br /><br /><br /><br />             
					 
					<label>Confirmed Date</label>
					<input type="text" name="Confirmed_Dt" id="Confirmed_Dt" class="date" value='<%Response.write rs("Confirmed_Dt") %>' />
					<label>Confirmed</label>
					<select name="Confirmed">
					<option value="<%= rs.Fields(43) %>"><%= rs.Fields(43) %></option>
					<option value="Y">Y</option>
					<option value="N">N</option>
					</select>                   
					<label>Admit Term</label>
					<input type="text" name="Admit_Term" id="Admit_Term" value='<%Response.write rs("Admit_Term") %>' readonly=true/>
					
					<br /><br /><br /><br />
					 <label>Withdraw Status</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;margin-right: 180px;" name="withdrawn" id="withdrawn" class="checkboxField" value='<%Response.write rs("withdrawn") %>' />
					<label>Withdraw Reason</label>
					<select  name="withdraw_reason" id="withdraw_reason">
					<option value="<%= rs.Fields(68) %>"><%= rs.Fields(68) %></option>
					<option value=" "> </option>
					<option value="WD">WD</option>
					<option value="AWD">AWD</option>
					<option value="WWD">WWD</option>
					<option value="S-AWD">S-AWD</option>
					<option value="S-SWD">S-SWD</option>
					<option value="S-WWD">S-WWD</option>
					</select>
					<label>Withdrawn Date</label>
					<input type="text" name="WithdrawnDate" id="WithdrawnDate"  class="date" value='<%Response.write rs("WithdrawnDate") %>' />
					
					
				   
					<br /><br /><br /><br />
					<label>Reapplicant</label>
					<select name="reapplicant">
					<option value="<%= rs.Fields(36) %>"><%= rs.Fields(36) %></option>                  
					<option value="Y">Y</option>
					<option value="N">N</option>
					</select>
					
					<label>Entered By</label>
					<input type="text" name="enteredby"  id="enteredby" value='<%Response.write rs("enteredby") %>'/>
					<label>Last Updated Date</label>
					<input type="text" name="LastUpdatedDt" id="LastUpdatedDt"  value='<%Response.write rs("LastUpdatedDt") %>' readonly=true/>

					<br /><br /><br /><br />
					<label>Limited Status</label>
					<select name="Limited_status">
					<option value='<%Response.write rs("Limited_status") %>'><%= rs.Fields(42) %></option>                  
					<option value="Y">Y</option>
					<option value="N">N</option>
					</select>
					<label>Received Deposit</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="received_deposit" id="received_deposit" class="checkboxField" value='<%Response.write rs("received_deposit") %>' />

					<label>Forward to Field</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="forward_to_field" id="forward_to_field" class="checkboxField" value='<%Response.write rs("forward_to_field") %>' />
					
					<br /><br /><br /><br />

					<label>Deferred From</label>
					<input type="text" name="DeferredFrom" id="DeferredFrom"  value='<%Response.write rs("DeferredFrom") %>'/>

					<label>Deferred To</label>
					<input type="text" name="DeferredTo" id="DeferredTo"  value='<%Response.write rs("DeferredTo") %>' />

                    <label>Adv Verification</label>
					<input type="checkbox" name="Adv_verification" style="margin:0;width:20px;height:20px;margin-right: 140px;" class="checkboxField" id="Adv_Verification"  value='<%Response.write rs("Adv_Verification") %>' />
					
					
					<br /><br /><br /><br />
					<label>Comments</label>
					<textarea id="comments" name="comments" cols="70" rows="5"><%Response.write rs("Comments") %></textarea> 
					</p>
					<br /><br /><br /><br />
					<p>
					
					<label>Credit In Statistics</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="Credit_in_Statistics" id="Credit_in_Statistics" class="checkboxField" value='<%Response.write rs("Credit_in_Statistics") %>'/>
					<label>Credit in BA BS</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="credit_in_ba_bs" id="credit_in_ba_bs" class="checkboxField" value='<%Response.write rs("credit_in_ba_bs") %>'/>
					<label>Credit in English</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="credit_in_english" id="credit_in_english" class="checkboxField" value='<%Response.write rs("credit_in_english") %>' />
					<br /><br /><br /><br />                    
					
					<label>Requesting Schools</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="requesting_schools" id="requesting_schools" class="checkboxField" value='<%Response.write rs("requesting_schools") %>' />
					<label>Financial Aid Request</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="Financial_Aid_Request" id="Financial_Aid_Request" class="checkboxField" value='<%Response.write rs("Financial_Aid_Request") %>' />
					<br/><br/><br/><br/>
					<label>Basic Skills TAP</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="Basic_Skill_Test" id="Basic_Skill_Test" class="checkboxField" value='<%Response.write rs("Basic_Skill_Test") %>' />
					<label>Basic Skills ACT/SAT</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="ACT_SAT" id="ACT_SAT" class="checkboxField" value='<%Response.write rs("ACT_SAT") %>' />
					<label>Passed Test</label>
					<select name="Passed_Test">
					<option value="<%Response.write rs("Passed_Test") %>"><%= rs.Fields(51) %></option>
					<option value="Y">Y</option>
					<option value="N">N</option>
					</select>
					<label>UG College</label>
					<select name="ugcollege" id="ugcollege">
					<option value="<%= rs.Fields(52) %>"><%= rs.Fields(52) %></option>    
					<option value="Colorado College">Colorado College</option>
					<option value="Hebrew University of Jerusalem">Hebrew University of Jerusalem</option>
					<option value="University of Iowa">University of Iowa</option>
					<option value="Addis Ababa University">Addis Ababa University</option>
					<option value="Adelphi University">Adelphi University</option>
					<option value="Adrian College">Adrian College</option>
					<option value="Ahmadu Bello University">Ahmadu Bello University</option>
					<option value="Alabama A & M University">Alabama A & M University</option>
					<option value="Alabama State University">Alabama State University</option>
					<option value="Albion College">Albion College</option>
					<option value="Alfred University">Alfred University</option>
					<option value="Allegheny College">Allegheny College</option>
					<option value="Alma College">Alma College</option>
					<option value="Alverno College">Alverno College</option>
					<option value="Ambrose University College">Ambrose University College</option>
					<option value="American Military University">American Military University</option>
					<option value="American Public University">American Public University</option>
					<option value="American University">American University</option>
					<option value="Anderson University">Anderson University</option>
					<option value="Andrews University">Andrews University</option>
					<option value="Ann Blitstein Institute">Ann Blitstein Institute</option>
					<option value="Antillian College">Antillian College</option>
					<option value="Antioch College">Antioch College</option>
					<option value="Appalachian State University">Appalachian State University</option>
					<option value="Aquinas College">Aquinas College</option>
					<option value="Arcadia University">Arcadia University</option>
					<option value="Argosy University">Argosy University</option>
					<option value="Arizona State University">Arizona State University</option>
					<option value="Arkansas State University">Arkansas State University</option>
					<option value="Armstrong Atlantic State">Armstrong Atlantic State</option>
					<option value="Art Institute of Chicago">Art Institute of Chicago</option>
					<option value="Asbury College">Asbury College</option>
					<option value="Asbury University">Asbury University</option>
					<option value="Ashford University">Ashford University</option>
					<option value="Ashland University">Ashland University</option>
					<option value="Associate Evangelical Seminary">Associate Evangelical Seminary</option>
					<option value="Athens State University">Athens State University</option>
					<option value="Auburn University">Auburn University</option>
					<option value="Augsburg College">Augsburg College</option>
					<option value="Augustana College">Augustana College</option>
					<option value="Auburn University">Auburn University</option>
					<option value="Aurora University">Aurora University</option>
					<option value="Austin College">Austin College</option>
					<option value="Austin Peay State University">Austin Peay State University</option>
					<option value="Azusa Pacific University">Azusa Pacific University</option>
					<option value="Babes-Bolyai University">Babes-Bolyai University</option>
					<option value="Baker College">Baker College</option>
					<option value="Baldwin-Wallace College">Baldwin-Wallace College</option>
					<option value="Ball State University">Ball State University</option>
					<option value="Bangalore University">Bangalore University</option>
					<option value="Barat College">Barat College</option>
					<option value="Bard College at Simons Rock College">Bard College at Simons Rock College</option>
					
					<option value="Bar-Ilan Univ.">Bar-Ilan Univ.</option>
					<option value="Bates College">Bates College</option>
					<option value="Baylor University">Baylor University</option>
					<option value="Beijing Institute of Fashion Technology">Beijing Institute of Fashion Technology</option>
					<option value="Beijing International Studies University">Beijing International Studies University</option>
					<option value="Beloit College">Beloit College</option>
					<option value="Bellevue University">Bellevue University</option>
					<option value="Belmont University">Belmont University</option>
					<option value="Benedict College">Benedict College</option>
					<option value="Benedictine University">Benedictine University</option>
					<option value="Bennett College">Bennett College</option>
					<option value="Bennington College">Bennington College</option>
					<option value="Berea College">Berea College</option>
					<option value="Bergen University College">Bergen University College</option>
					<option value="Bethel College">Bethel College</option>
					<option value="Bethel University">Bethel University</option>
					<option value="Bharathiar University">Bharathiar University</option>
					<option value="Bhavnagar University">Bhavnagar University</option>
					
					<option value="Binghamton University">Binghamton University</option>
					<option value="Biola University">Biola University</option>
					<option value="Birzeit University">Birzeit University</option>
					<option value="Blackburn College">Blackburn College</option>
					<option value="Blitstein Institute of Hebrew Theological College">Blitstein Institute of Hebrew Theological College</option>
					<option value="Bloomsburg University">Bloomsburg University</option>
					<option value="Bob Jones University">Bob Jones University</option>
					<option value="Bogazici University">Bogazici University</option>
					<option value="Boise State University">Boise State University</option>
					<option value="Boston College">Boston College</option>
					<option value="Boston Conservatory">Boston Conservatory</option>
					<option value="Boston University">Boston University</option>
					<option value="Bowdoin College">Bowdoin College</option>
					<option value="Bowling Green State University">Bowling Green State University</option>
					<option value="Bradley University">Bradley University</option>
					<option value="Brandeis University">Brandeis University</option>
					<option value="Briar Cliff College">Briar Cliff College</option>
					<option value="Brigham Young University">Brigham Young University</option>
					<option value="Brigham Young University Hawaii">Brigham Young University Hawaii</option>
			<option value="Brigham Young University of Idaho">Brigham Young University of Idaho</option>
			
					<option value="Brown University">Brown University</option>
					<option value="Bryant University">Bryant University</option>
					<option value="Buena Vista University">Buena Vista University</option>
					<option value="Butler University">Butler University</option>
					<option value="BYU">BYU</option>
					<option value="California Baptist University">California Baptist University</option>
					<option value="California Lutheran University">California Lutheran University</option>
					<option value="California Polytechnic State University">California Polytechnic State University</option>
					<option value="California State University">California State University</option>
					<option value="California State University Chico">California State University Chico</option>
					<option value="California State University Dominguez Hills">California State University Dominguez Hills</option>
					<option value="California State University East Bay">California State University East Bay</option>
					<option value="California State University Fresno">California State University Fresno</option>
					<option value="California State University Fullerton">California State University Fullerton</option>
					<option value="California State University Long Beach">California State University Long Beach</option>
					<option value="California State University Los Angeles">California State University Los Angeles</option>
					<option value="California State University Northridge">California State University Northridge</option>
					<option value="California State University Sacramento">California State University Sacramento</option>
					<option value="California State University San Bernardino">California State University San Bernardino</option>
					<option value="California University of Pennsylvania">California University of Pennsylvania</option>
					<option value="Calumet College of St. Joseph">Calumet College of St. Joseph</option>
					<option value="Calvin College">Calvin College</option>
					<option value="Campbell University">Campbell University</option>
					<option value="Campbellsville University">Campbellsville University</option>
				
					<option value="Canisius College">Canisius College</option>
					<option value="Capitol Normal University, Beijing">Capitol Normal University, Beijing</option>
					<option value="Cardinal Stritch University">Cardinal Stritch University</option>
					<option value="Carleton College">Carleton College</option>
					<option value="Carleton University">Carleton University</option>
					<option value="Carnegie Mellon University">Carnegie Mellon University</option>
					<option value="Carroll College">Carroll College</option>
					<option value="Carroll University">Carroll University</option>
					<option value="Carson Newman College">Carson Newman College</option>
					<option value="Carthage">Carthage</option>
					<option value="Carthage College">Carthage College</option>
					<option value="Case Western Reserve University">Case Western Reserve University</option>
					<option value="Catholic University">Catholic University</option>
					<option value="Cedarville College">Cedarville College</option>
					<option value="Cedarville University">Cedarville University</option>
					<option value="Centenary College">Centenary College</option>
					<option value="Central  Michigan University">Central  Michigan University</option>
					<option value="Central College">Central College</option>
					<option value="Central State University">Central State University</option>
					<option value="Central University of Finance and Econ">Central University of Finance and Econ</option>
					<option value="Centre College">Centre College</option>
					<option value="Changsha Medical University">Changsha Medical University</option>
					<option value="Charles University of Prague">Charles University of Prague</option>
					<option value="Chatham University">Chatham University</option>
					<option value="Chicago State University">Chicago State University</option>
					<option value="Chico State University">Chico State University</option>
					<option value="China Womens University">China Womens University</option>
					<option value="Chinese University Hong Kong">Chinese University Hong Kong</option>
					<option value="Chonbuk National University">Chonbuk National University</option>
					<option value="Christopher Newport University">Christopher Newport University</option>
					<option value="Chung-Ang University">Chung-Ang University</option>
					<option value="Chungnam National University">Chungnam National University</option>
					<option value="Chuo Un.">Chuo Un.</option>
					<option value="City College of New York">City College of New York</option>
					<option value="City University of Hong Kong">City University of Hong Kong</option>
					<option value="Claremont McKenna College">Claremont McKenna College</option>
					<option value="Clarion University">Clarion University</option>
					<option value="Clark Atlanta University">Clark Atlanta University</option>
					<option value="Clarke College">Clarke College</option>
					<option value="Clarke University">Clarke University</option>
					<option value="Clearwater Christian College-Florida">Clearwater Christian College-Florida</option>
					<option value="Cleary University">Cleary University</option>
					<option value="Clemson University">Clemson University</option>
					<option value="Cleveland State University">Cleveland State University</option>
					<option value="Coastal Carolina University">Coastal Carolina University</option>
					<option value="Coe College">Coe College</option>
					<option value="Colby College">Colby College</option>
					<option value="Colgate University">Colgate University</option>
					<option value="College of Charleston">College of Charleston</option>
					<option value="College of DuPage">College of DuPage</option>
					<option value="College of DePaul">College of DePaul</option>
					<option value="College of Lake Co.">College of Lake Co.</option>
					<option value="College of Lake County">College of Lake County</option>
					<option value="College of Mt. St. Joseph">College of Mt. St. Joseph</option>
					<option value="College of Ozarks">College of Ozarks</option>
					<option value="College of Saint Benedict/St. Johns University">College of Saint Benedict/St. Johns University</option>
					<option value="College of Saint Benedicts">College of Saint Benedicts</option>
					<option value="College of Santa Fe">College of Santa Fe</option>
					<option value="College of St Benedict">College of St Benedict</option>
					<option value="College of St Catherine">College of St Catherine</option>
					<option value="College of St Scholastica">College of St Scholastica</option>
					
					<option value="College of St. Francis">College of St. Francis</option>
					<option value="College of the Holy Cross">College of the Holy Cross</option>
					<option value="College of the Ozarks">College of the Ozarks</option>
					<option value="College of William and Mary">College of William and Mary</option>
					<option value="College of Wooster">College of Wooster</option>
					<option value="Colorado Christian University">Colorado Christian University</option>
					<option value="Colorado College">Colorado College</option>
					<option value="Colorado Mesa State">Colorado Mesa State</option>
					<option value="Colorado State University">Colorado State University</option>
					<option value="Columbia College - Missouri">Columbia College - Missouri</option>
					<option value="Columbia College Chicago">Columbia College Chicago</option>
					<option value="Columbia University">Columbia University</option>
					<option value="Concordia University">Concordia University</option>
					<option value="Concordia River Forest">Concordia River Forest</option>
					<option value="Concordia University">Concordia University</option>
					<option value="Concordia University River Forest">Concordia University River Forest</option>
					<option value="Concordia University Wisconsin">Concordia University Wisconsin</option>
					<option value="Connecticut College">Connecticut College</option>
					<option value="Cornell College">Cornell College</option>
					<option value="Cornell University">Cornell University</option>
					<option value="Cornerstone University">Cornerstone University</option>
					<option value="Costo Carolina University">Costo Carolina University</option>
					<option value="Covenant College">Covenant College</option>
					<option value="Creighton University">Creighton University</option>
					<option value="Chung-Ang University">Chung-Ang University</option>
					<option value="CUNY John Jay Coll Crimnl Just">CUNY John Jay Coll Crimnl Just</option>
					<option value="CUNY-Buffalo">CUNY-Buffalo</option>
					<option value="Daemen College">Daemen College</option>
					<option value="Dakota State University">Dakota State University</option>
					<option value="Dalhousie University">Dalhousie University</option>
					<option value="Dalian Maritime University">Dalian Maritime University</option>
					<option value="Dalton State College">Dalton State College</option>
					<option value="Dana College">Dana College</option>
					<option value="Dartmouth College">Dartmouth College</option>
					<option value="De Paul University">De Paul University</option>
					<option value="Defiance College">Defiance College</option>
					<option value="Delaware State University">Delaware State University</option>
					<option value="Delta State University">Delta State University</option>
					<option value="Denison University">Denison University</option>
					<option value="DePauw University">DePauw University</option>
					<option value="Devi Ahilya University Idore (Madhya Pradesh, India)">Devi Ahilya University Idore (Madhya Pradesh, India)</option>
					<option value="Devry University">Devry University</option>
					<option value="Dillard University">Dillard University</option>
					<option value="Dominican University">Dominican University</option>
					<option value="Dongguk University">Dongguk University</option>
					<option value="Dordt College">Dordt College</option>
					<option value="Dowling College">Dowling College</option>
					<option value="Drake University">Drake University</option>
					<option value="Drew University">Drew University</option>
					<option value="Duke University">Duke University</option>
					<option value="Duquesne University">Duquesne University</option>
					<option value="Earlham College">Earlham College</option>
					<option value="East China Normal University">East China Normal University</option>
					<option value="East Tennessee State">East Tennessee State</option>
					<option value="East West University">East West University</option>
					<option value="Eastern Illinois University">Eastern Illinois  University</option>
					<option value="Eastern Kentucky University">Eastern Kentucky University</option>
					<option value="Eastern Mennonite University">Eastern Mennonite University</option>
					<option value="Eastern Michigan University">Eastern Michigan University</option>
					<option value="Eastern New Mexico">Eastern New Mexico</option>
			<option value="Eastern Oregon University">Eastern Oregon University</option>
					
					<option value="Eastern Stroudsburg University">Eastern Stroudsburg Univiversity</option>
					<option value="Eastern washington University">Eastern washington University</option>
					<option value="East-West University">East-West University</option>
					<option value="Eckerd College">Eckerd College</option>
					<option value="Edgewood College">Edgewood College</option>
					<option value="Edinboro University">Edinboro University</option>
					
					<option value="Elgin Community College">Elgin Community College</option>
					<option value="Elhurst College">Elhurst College</option>
					<option value="Elizabethtown College">Elizabethtown College</option>
					<option value="Elmhurst College">Elmhurst College</option>
					<option value="Elon College">Elon College</option>
					<option value="Elon University">Elon University</option>
					<option value="Emerson College">Emerson College</option>
					<option value="Emmanuel College">Emmanuel College</option>
					<option value="Emmaus Bible College">Emmaus Bible College</option>
					<option value="Emory University">Emory University</option>
					<option value="Emporia State University">Emporia State University</option>
					<option value="ESUT Business School">ESUT Business School</option>
					<option value="Ethiraj College for Women">Ethiraj College for Women</option>
					<option value="Eugene Lang College/New School University">Eugene Lang College/New School University</option>
					<option value="Evangel University">Evangel University</option>
					<option value="Evergreen State College">Evergreen State College</option>
					<option value="Evergreen State University">Evergreen State University</option>
					<option value="Ewha Womans University">Ewha Womans University</option>
					
					<option value="Excelsior College">Excelsior College</option>
					<option value="Fairfield University">Fairfield University</option>
					<option value="Ferris State University">Ferris State University</option>
					<option value="Findlay University">Findlay University</option>
					<option value="Fisk University">Fisk University</option>
					<option value="Flagler College">Flagler College</option>
					<option value="Florida A&M University">Florida A&M University</option>
					<option value="Florida Atlantic University">Florida Atlantic University</option>
					<option value="Florida Community College">Florida Community College</option>
					<option value="Florida Gulf Coast">Florida Gulf Coast</option>
					<option value="Florida Gulf Coast University">Florida Gulf Coast University</option>
					<option value="Florida International University">Florida International University</option>
					<option value="Florida State University">Florida State University</option>
					<option value="Fontbonne University">Fontbonne University</option>
					<option value="Fordham University">Fordham University</option>
					<option value="Fort Lewis College">Fort Lewis College</option>
					<option value="Fort Valley State University">Fort Valley State University</option>
					<option value="Fourah Bay College">Fourah Bay College</option>
					<option value="Franciscan University">Franciscan University</option>
					<option value="Franciscan University of Steubenville">Franciscan University of Steubenville</option>
					<option value="Franklin and Marshall College">Franklin and Marshall College</option>
					<option value="Ft. Scott Community College">Ft. Scott Community College</option>
					<option value="Ft. Valley State University">Ft. Valley State University</option>
					<option value="Fudan University">Fudan University</option>
					<option value="Fullerton College">Fullerton College</option>
					<option value="Furman University">Furman University</option>
					
					<option value="George Mason University">George Mason University</option>
					<option value="George Washington University">George Washington University</option>
					<option value="George Williams College at Au">George Williams College at Au</option>
					<option value="George Williams University">George Williams University</option>
					<option value="Georgetown University">Georgetown University</option>
					<option value="Georgia State University">Georgia State University</option>
					<option value="George Washington University">George Washington University</option>
					<option value="George Fox University">George Fox University</option>
					<option value="Georgetown University">Georgetown University</option>
					<option value="Gettysburg College">Gettysburg College</option>
					
					<option value="Gonzaga University">Gonzaga University</option>
					<option value="Gordon College">Gordon College</option>
					<option value="Goshen College">Goshen College</option>
					<option value="Goucher College">Goucher College</option>e
					<option value="Governors State University">Governors State University</option>
					<option value="Grambling State University">Grambling State University</option>
					<option value="Grand Valley State University">Grand Valley State University</option>
					<option value="Green Mountain College">Green Mountain College</option>
					<option value="Greenville College">Greenville College</option>
					<option value="Grinnell College">Grinnell College</option>
					
					<option value="Guangzhou Normal University">Guangzhou Normal University</option>
					<option value="Guilford College">Guilford College</option>
					<option value="Gustavus Adolphus College">Gustavus Adolphus College</option>
					<option value="Hamilton College">Hamilton College</option>
					<option value="Hamline University">Hamline University</option>
					<option value="Hampshire College">Hampshire College</option>
					<option value="Hampton Institute">Hampton Institute</option>
					<option value="Hampton University">Hampton University</option>
					<option value="Han Shin University">Han Shin University</option>
					<option value="Hankuk University of Foreign Studies">Hankuk University of Foreign Studies</option>
					<option value="Hannam University">Hannam University</option>
					<option value="Hansung Univ">Hansung Univ</option>
					<option value="Hanyang University">Hanyang University</option>
					<option value="Harding University">Harding University</option>
					<option value="Harold Washington College">Harold Washington College</option>
					<option value="Harvard College">Harvard College</option>
					<option value="Harvard University">Harvard University</option>
					<option value="Haverford College">Haverford College</option>
					<option value="Hawaii Pacific University">Hawaii Pacific University</option>
					<option value="Hebrew Theological College">Hebrew Theological College</option>
					<option value="Hebrew University">Hebrew University</option>
					<option value="Hebrew University of Jerusalem">Hebrew University of Jerusalem</option>
					<option value="Heidelberg College">Heidelberg College</option>
					<option value="Hendrix College">Hendrix College</option>
					<option value="Hewbrew Theological College for Women">Hewbrew Theological College for Women</option>
					<option value="Hillsdale College">Hillsdale College</option>
					<option value="Hiram College">Hiram College</option>
					<option value="Hiroshima Prefectural Womens University">Hiroshima Prefectural Womens University</option>
					<option value="Hiroshima Shudo University">Hiroshima Shudo University</option>
					<option value="Hofstra University">Hofstra University</option>
					<option value="Hollins University">Hollins University</option>
					<option value="Hong Kong Baptist University">Hong Kong Baptist University</option>
					<option value="Hope College">Hope College</option>
					<option value="Houghton College">Houghton College</option>
					<option value="Howard University">Howard University</option>
					<option value="Hubei Engineering University">Hubei Engineering University</option>
					<option value="Hunter College">Hunter College</option>
					<option value="Huntington University">Huntington University</option>
					<option value="IA State Univ.">IA State Univ.</option>
					<option value="IIT">IIT</option>
					<option value="Illinois Wesleyan University">Illinois Wesleyan University</option>
					
					<option value="Illinois Benedctine College">Illinois Benedctine College</option>
					<option value="Illinois Central College">Illinois Central College</option>
					<option value="Illinois College">Illinois College</option>
					<option value="Illinois Institute of Art">Illinois Institute of Art</option>
					<option value="Illinois State University">Illinois State University</option>
					<option value="Indiana State University">Indiana State University</option>
					<option value="Indiana University">Indiana University</option>
					<option value="Indiana University Purdue">Indiana University Purdue</option>
					<option value="Indiana University Bloomington">Indiana University Bloomington</option>
					<option value="Indiana University Calumet">Indiana University Calumet</option>
					<option value="Indiana University Indianapolis">Indiana University Indianapolis</option>
					<option value="Indiana University Northwest">Indiana University Northwest</option>
					<option value="Indiana University of Pennsylvania">Indiana University of Pennsylvania</option>
					<option value="Indiana University Pennsylvania">Indiana University Pennsylvania</option>
					<option value="Indiana University South Bend">Indiana University South Bend</option>
					<option value="Indiana Wesleyan University">Indiana Wesleyan University</option>
					<option value="Indraprastha College">Indraprastha College</option>
					<option value="Institute for Tourism Studies">Institute for Tourism Studies</option>
					<option value="Inter American University">Inter American University</option>
					<option value="International Academy of Design & Technology">International Academy of Design & Technology</option>
					<option value="Iona College">Iona College</option>
					<option value="Iowa State University">Iowa State University</option>
					<option value="Islamic Azad University">Islamic Azad University</option>
					<option value="Istanbul University">Istanbul University</option>
					<option value="ISWR, Dhaka University">ISWR, Dhaka University</option>
					<option value="Ithaca College">Ithaca College</option>
					<option value="Jackson State University">Jackson State University</option>
					<option value="Jacksonville University">Jacksonville University</option>
					<option value="Jagiellonian University">Jagiellonian University</option>
					<option value="James Madison University">James Madison University</option>
					<option value="Jawaharlal Nehru University">Jawaharlal Nehru University</option>
					<option value="Jilin University">Jilin University</option>
					<option value="John Carroll University">John Carroll University</option>
					<option value="John Jay College of Criminal Justice">John Jay College of Criminal Justice</option>
					<option value="Johns Hopkins University">Johns Hopkins University</option>
					<option value="Joliet College">Joliet College</option>
					
					<option value="Judson University">Judson University</option>
					<option value="Juniata College">Juniata College</option>
					<option value="Kalamazoo College">Kalamazoo College</option>
					<option value="Kangnam University">Kangnam University</option>
					<option value="Kankakee Community College">Kankakee Community College</option>
					<option value="Kansai Gaidai College">Kansai Gaidai College</option>
					<option value="Kansas State University">Kansas State University</option>
					<option value="Kanto Gakuin University">Kanto Gakuin University</option>
					<option value="Kaplan University">Kaplan University</option>
					<option value="Keene State College">Keene State College</option>
					<option value="Kendall College">Kendall College</option>
					<option value="Kennedy King Community College">Kennedy King Community College</option>
					<option value="Kennesaw State University">Kennesaw State University</option>
					<option value="Kent State University">Kent State University</option>
					<option value="Kentucky Christian College">Kentucky Christian College</option>
					<option value="Kentucky Wesleyan College">Kentucky Wesleyan College</option>
					<option value="Kenyon College">Kenyon College</option>
					<option value="Keuka College">Keuka College</option>
					<option value="Kkottongnae Hyundo University- Korea">Kkottongnae Hyundo University- Korea</option>
					<option value="Knox College">Knox College</option>
					<option value="Koc University">Koc University</option>
					<option value="Kon-Kuk U Seoul">Kon-Kuk U Seoul</option>
					<option value="Korea University">Korea University</option>
					<option value="Kumamoto Gakuen University">Kumamoto Gakuen University</option>
					<option value="Kuyper College">Kuyper College</option>
					<option value="Kyungsung University">Kyungsung University</option>
					<option value="La Grange Community College">La Grange Community College</option>
					<option value="Lake Forest College">Lake Forest College</option>
					<option value="Lake Superior State University">Lake Superior State University</option>
					<option value="Lakeland College">Lakeland College</option>
					<option value="Lamar University">Lamar University</option>
					<option value="Lander University">Lander University</option>
					<option value="Lane College">Lane College</option>
					<option value="Langston University">Langston University</option>
					<option value="La Salle University">La Salle University</option>
					<option value="Lawrence University">Lawrence University</option>
					<option value="Lee University">Lee University</option>
					<option value="Lehigh University">Lehigh University</option>
					<option value="LeMoyne College">LeMoyne College</option>
					<option value="Lenoir Rhyne College">Lenoir Rhyne College</option>
					<option value="Lesley College">Lesley College</option>
					<option value="Lewis & Clark University">Lewis & Clark University</option>
					<option value="Lewis University">Lewis University</option>
					<option value="Liberty University">Liberty University</option>
					<option value="Limestone College">Limestone College</option>
					<option value="Lincoln Christian College">Lincoln Christian College</option>
					<option value="Lincoln University">Lincoln University</option>
					<option value="Lindenwood University">Lindenwood University</option>
					<option value="Lindenwood University-Belleville">Lindenwood University-Belleville</option>
					<option value="Lindsey Wilson College">Lindsey Wilson College</option>
					<option value="Linkoping University">Linkoping University</option>
					<option value="Lipscomb University">Lipscomb University</option>
					<option value="Livingstone College">Livingstone College</option>
					<option value="Lock Haven University of PA">Lock Haven University of PA</option>
					<option value="Long Island University">Long Island University</option>
					<option value="loras College">loras College</option>
					<option value="Lords College">Lords College</option>
					<option value="Louisiana State University">Louisiana State University</option>
					<option value="Louisiana Tech University">Louisiana Tech University</option>
					<option value="Loyola College of Maryland">Loyola College of Maryland</option>
					<option value="Loyola Marymount University">Loyola Marymount University</option>
					<option value="Loyola University Chicago">Loyola University Chicago</option>
					<option value="Loyola University New Orleans">Loyola University New Orleans</option>
					<option value="Loyola University of Maryland">Loyola University of Maryland</option>
					<option value="Luther College">Luther College</option>
					<option value="Lycoming College">Lycoming College</option>
					<option value="Lynn University">Lynn University</option>
					<option value="Mac Alester College">Mac Alester College</option>
					<option value="Mac Murray College">Mac Murray College</option>
					<option value="Macalester College">Macalester College</option>
					<option value="MacMurray College">MacMurray College</option>
					<option value="Macquarie University">Macquarie University</option>
					<option value="Madonna University">Madonna University</option>
					<option value="Madras University">Madras University</option>
					<option value="Mahatma Gandhi University">Mahatma Gandhi University</option>
					<option value="Malone College">Malone College</option>
					<option value="Manchester College">Manchester College</option>
					<option value="Maranatha Baptist Bible College">Maranatha Baptist Bible College</option>
					<option value="Marian College Indianapolis">Marian College Indianapolis</option>
					<option value="Marian College of Fond Du Lac">Marian College of Fond Du Lac</option>
					<option value="Marietta College">Marietta College</option>
					<option value="Marist College">Marist College</option>
					<option value="Marquette University">Marquette University</option>
					<option value="Mary Washington College">Mary Washington College</option>
					<option value="Marygrove College">Marygrove College</option>
					<option value="Maryland Institute College of Art">Maryland Institute College of Art</option>
					<option value="Marymount Manhattan college">Marymount Manhattan college</option>
					<option value="Maryville University">Maryville University</option>
					<option value="Maryville University of St Louis">Maryville University of St Louis</option>
					<option value="Marywood University">Marywood University</option>
					<option value="Massey University">Massey University</option>
					<option value="McGill University">McGill University</option>
					<option value="McHenry County College">McHenry County College</option>
					<option value="McKendree University">McKendree University</option>
					<option value="Mercy College">Mercy College</option>
					<option value="Mercyhurst College">Mercyhurst College</option>
					<option value="Messiah College">Messiah College</option>
					<option value="Methodist College">Methodist College</option>
					<option value="Metro State College">Metro State College</option>
					<option value="Metropolitan Autonomous University">Metropolitan Autonomous University</option>
					<option value="Metropolitan State College of Denver">Metropolitan State College of Denver</option>
					<option value="Metropolitan State University">Metropolitan State University</option>
					<option value="Miami University">Miami University</option>
					<option value="Miami University-Oxford Ohio">Miami University-Oxford Ohio</option>
					<option value="Miami of Ohio University">Miami of Ohio University</option>
					<option value="Michigan State University">Michigan State University</option>
					<option value="MidAmerica Bible College">MidAmerica Bible College</option>
					<option value="Middle Tennessee State University">Middle Tennessee State University</option>
					<option value="Midwestern State University">Midwestern State University</option>
					<option value="Millikin University">Millikin University</option>
					<option value="Millersville University of Pennsylvania">Millersville University of Pennsylvania</option>
					<option value="Milligan College">Milligan College</option>
					<option value="Minnesota State University">Minnesota State University</option>
					<option value="Minnesota State University-Mankato">Minnesota State University-Mankato</option>
					<option value="Mississippi State University">Mississippi State University</option>
					<option value="Mississippi Valley State">Mississippi Valley State</option>
					<option value="Missouri Baptist University">Missouri Baptist University</option>
					<option value="Missouri State University">Missouri State University</option>
					<option value="Missouri Valley College">Missouri Valley College</option>
					<option value="Missouri Western State University">Missouri Western State University</option>
					<option value="Moddy Bible Institute">Moddy Bible Institute</option>
					<option value="Monmouth College">Monmouth College</option>
					<option value="Montana State University">Montana State University</option>
					<option value="Montana State University Bozeman">Montanta State University Bozeman</option>
					<option value="Montclair University">Montclair University</option>
					<option value="Moody Bible Institute">Moody Bible Institute</option>
					<option value="Moorhead State University">Moorhead State University</option>
					<option value="Morehouse College">Morehouse College</option>
					<option value="Morgan State University">Morgan State University</option>
					<option value="Moraine Valley Community College">Moraine Valley Community College</option>
					<option value="Morris Brown College">Morris Brown College</option>
					<option value="Moscow Lomonosov State University">Moscow Lomonosov State University</option>
					<option value="Mother Pattern College of Health Sciences">Mother Pattern College of Health Sciences</option>
					<option value="Mount Holyoke College">Mount Holyoke College</option>
					<option value="Mount St. Clare">Mount St. Clare</option>
					<option value="Mt. Holyoke">Mt. Holyoke</option>
					<option value="Mt. Mercy College">Mt. Mercy College</option>
					
					<option value="Mt. Saint Marys College">Mt. Saint Marys College</option>
					<option value="Mt. Vernon Nazarene University">Mt. Vernon Nazarene University</option>
					<option value="Muhlenberg College">Muhlenberg College</option>
					<option value="Mumbai University">Mumbai University</option>
					<option value="Mundelein College">Mundelein College</option>
					<option value="Murray State University">Murray State University</option>
					<option value="NAES College">NAES College</option>
					<option value="Nanjing University Science and Tech">Nanjing University Science and Tech</option>
					<option value="Naropa University">Naropa University</option>
					<option value="National Chengchi University">National Chengchi University</option>
					<option value="National Chi-Nan University -Taiwan-Rep of China">National Chi-Nan University -Taiwan-Rep of China</option>
					<option value="National Chung Cheng University">National Chung Cheng University</option>
					<option value="National Louis University">National Louis University</option>
					<option value="National Taiwan University">National Taiwan University</option>
				<option value="National University of Asuncion">National University of Asuncion</option>
					<option value="National University of Ho Chi Minh City">National University of Ho Chi Minh City</option>
					<option value="Nazareth College">Nazareth College</option>
					<option value="Nazareth College of Rochester">Nazareth College of Rochester</option>
					<option value="Nebraska Wesleyan University">Nebraska Wesleyan University</option>
					<option value="New College of California">New College of California</option>
					<option value="New College of University of South Florida">New College of University of South Florida</option>
					<option value="New Mexico State University">New Mexico State University</option>
					<option value="New School University">New School University</option>
					<option value="New York University">New York University</option>
					<option value="Nihon University">Nihon University</option>
					<option value="Norfolk State University">Norfolk State University</option>
					<option value="North Carolina A & T University">North Carolina A & T University</option>
					<option value="North Carolina State University">North Carolina State University</option>
					<option value="North Central College">North Central College</option>
					<option value="North Central University">North Central University</option>
					<option value="North Dakota State University">North Dakota State University</option>
					<option value="North Kazakhstan State University">North Kazakhstan State University</option>
					<option value="North Park University">North Park University</option>
					<option value="North Park College">North Park College</option>
					<option value="Northern Arizona University">Northern Arizona University</option>
					<option value="Northeastern Illinois University">Northeastern Illinois University</option>
					<option value="Northeastern University- Boston">Northeastern University- Boston</option>
					<option value="Northern Michigan University">Northern Michigan University</option>
					<option value="Northern Arizona State University">Northern Arizona State University</option>
					<option value="Northern Illinois University">Northern Illinois University</option>
					<option value="Northern Kentucky University">Northern Kentucky University</option>
					<option value="Northern Michigan University">Northern Michigan University</option>
					<option value="Northland College">Northland College</option>
					<option value="Northland International University">Northland International University</option>
					<option value="Northwest University">Northwest University</option>
					<option value="Northwestern University">Northwestern University</option>
					<option value="Northwestern College">Northwestern College</option>
					<option value="Northwestern College Iowa">Northwestern College Iowa</option>
					<option value="Northwestern College Minnesota">Northwestern College Minnesota</option>
					<option value="Northwestern University Bienen School of Music">Northwestern University Bienen School of Music</option>
				   
					<option value="Notre Dame University">Notre Dame University</option>
					<option value="Nottingham Trent University">Nottingham Trent University</option>
					<option value="Oakland University">Oakland University</option>
					<option value="Oakton Community College">Oakton Community College</option>
					<option value="Oakwood College">Oakwood College</option>
					<option value="Oberlin College">Oberlin College</option>
					<option value="Oberlin Conservatory of Music">Oberlin Conservatory of Music</option>
					<option value="Occidental College">Occidental College</option>
					<option value="Ohio Northern University">Ohio Northern University</option>
					<option value="Ohio State University">Ohio State University</option>
					<option value="Ohio State university Columbus">Ohio State university Columbus</option>
					<option value="Ohio Wesleyan University">Ohio Wesleyan University</option>
					<option value="Oklahoma Baptist University">Oklahoma Baptist University</option>
					<option value="Oklahoma City University">Oklahoma City University</option>
					<option value="Oklahoma State University">Oklahoma State University</option>
					<option value="Oklahoma State University-Stillwater">Oklahoma State University-Stillwater</option>
					<option value="Old Dominion University">Old Dominion University</option>
					<option value="Olive Harvey Community College">Olive Harvey Community College</option>
					<option value="Olivet College">Olivet College</option>
					<option value="Olivet Nazarene University">Olivet Nazarene University</option>
					<option value="Oral Roberts University">Oral Roberts University</option>
					<option value="Oregon State University">Oregon State University</option>
					<option value="Osmania University">Osmania University</option>
					<option value="Ottawa University">Ottawa University</option>
					<option value="Pace University">Pace University</option>
					<option value="Pacific Lutheran University">Pacific Lutheran University</option>
					<option value="Palm Beach Atlantic University">Palm Beach Atlantic University</option>
					<option value="Park University">Park University</option>
					<option value="Parkland College">Parkland College</option>
					<option value="Peking University">Peking University</option>
					<option value="Pennsylvania State University">Pennsylvania State University</option>
					<option value="Pepperdine University">Pepperdine University</option>
					<option value="Philander Smith College">Philander Smith College</option>
					<option value="Pitzer College">Pitzer College</option>
					<option value="Point Loma Nazarene University">Point Loma Nazarene University</option>
					<option value="Point Park College">Point Park College</option>
					<option value="Pomona College">Pomona College</option>
					<option value="Portland State University">Portland State University</option>
					<option value="Prairie State College">Prairie State College</option>
					<option value="Prairie View A & M University">Prairie View A & M University</option>
					<option value="Princeton University">Princeton University</option>
					<option value="Providence College">Providence College</option>
					<option value="Purchase College, SUNY">Purchase College, SUNY</option>
					<option value="Purdue University">Purdue University</option>
					<option value="Purdue University Calumet">Purdue University Calumet</option>
					<option value="Purdue University Northwest">Purdue University Northwest</option>
					<option value="Qingdao Binhai University">Qingdao Binhai University</option>
					<option value="Queens College CUNY">Queens College CUNY</option>
					<option value="Queens University">Queens University</option>
					<option value="Quincy University">Quincy University</option>
					<option value="Radcliffe College">Radcliffe College</option>
					<option value="Radford University">Radford University</option>
					<option value="Randolph-Macon College">Randolph-Macon College</option>
					<option value="Redeemer University">Redeemer University</option>
					<option value="Reed College">Reed College</option>
					<option value="Reformed Bible College">Reformed Bible College</option>
					<option value="Regis College">Regis College</option>
					<option value="Regis University">Regis University</option>
					<option value="Renmin University">Renmin University</option>
					<option value="Rensselaer Polytechnic Institute">Rensselaer Polytechnic Institute</option>
					<option value="Rhode Island College">Rhode Island College</option>
					<option value="Rhodes College">Rhodes College</option>
					<option value="Rice University">Rice University</option>
					<option value="Richard Stockton College of New Jersey">Richard Stockton College of New Jersey</option>
					<option value="Richmond College">Richmond College</option>
					<option value="Rider University">Rider University</option>
					<option value="Ripon College">Ripon College</option>
					<option value="Robert Morris University Chicago">Robert Morris University Chicago</option>
					<option value="Robert Morris University Pennsylvania">Robert Morris University Pennsylvania</option>
					<option value="Roberts Wesleyan College">Roberts Wesleyan College</option>
					<option value="Rochester Institute of Technology">Rochester Institute of Technology</option>
					<option value="Rockford Community College">Rockford Community College</option>
					<option value="Rocky Mountain College">Rocky Mountain College</option>
					<option value="Rollins College">Rollins College</option>
					<option value="Roosevelt University">Roosevelt University</option>
					<option value="Rosary College">Rosary College</option>
					<option value="Rose Hulman Inst of Technology">Rose Hulman Inst of Technology</option>
					<option value="Rowan University">Rowan University</option>
					<option value="Rush University">Rush University</option>
					<option value="Rust College">Rust College</option>
					<option value="Russell Sage College">Russell Sage College</option>
					<option value="Rutgers University">Rutgers University</option>
					<option value="Rutgers University New Brunswick">Rutgers University New Brunswick</option>
					<option value="Sage Colleges">Sage Colleges</option>
					<option value="Saginaw Valley State University">Saginaw Valley State University</option>
					<option value="Sahmyook University">Sahmyook University</option>
					<option value="Saint Anselm College">Saint Anselm College</option>
					<option value="Saint Leo University">Saint Leo University</option>
					<option value="Saint Louis University">Saint Louis University</option>
					<option value="Saint Marys College">Saint Marys College</option>
					<option value="Saint Marys University - Minnesota">Saint Marys University - Minnesota</option>
					<option value="Saint Norbert College">Saint Norbert College</option>
					<option value="Saint Xavier University">Saint Xavier University</option>
					<option value="Saint Marys College of Notre Dame">Saints Mary College of Notre Dame</option>
					<option value="Salford University">Salford University</option>
					<option value="Salve Regina University">Salve Regina University</option>
					<option value="Sam Houston State University">Sam Houston State University</option>
					<option value="San Antonio College">San Antonio College</option>
					<option value="San Diego State University">San Diego State University</option>
					<option value="San Francisco State University">San Francisco State University</option>
					<option value="San Jose State University">San Jose State University</option>
					<option value="Santa Clara University">Santa Clara University</option>
					<option value="Sarah Lawrence College">Sarah Lawrence College</option>
					<option value="School of the Art Institute">School of the Art Institute</option>
					<option value="Scripps College">Scripps College</option>
					<option value="Seton Hall University">Seton Hall University</option>
					<option value="Seattle University">Seattle University</option>
					<option value="Seoul Jangsin University">Seoul Jangsin University</option>
					<option value="Seton Hall University">Seton Hall University</option>
					<option value="Sewanee: The University of The South">Sewanee: The University of The South</option>
					<option value="Shaw University">Shaw University</option>
					<option value="Shawnee State University">Shawnee State University</option>
					<option value="Shenzhen University">Shenzhen University</option>
					<option value="Shih Chien University">Shih Chien University</option>
					<option value="Shimer College">Shimer College</option>
					<option value="Shor Yoshuv Rabbinical College">Shor Yoshuv Rabbinical College</option>
					<option value="Siena Heights University">Siena Heights University</option>
					<option value="Simmons College">Simmons College</option>
					<option value="Simpson College">Simpson College</option>
					<option value="Simpson University">Simpson University</option>
					<option value="Skidmore College">Skidmore College</option>
					<option value="Smith College">Smith College</option>
					<option value="Sonoma State University">Sonoma State University</option>
					<option value="Soochow University">Soochow University</option>
					<option value="Soongsil University">Soongsil University</option>
					<option value="Sophia University">Sophia University</option>
					<option value="South Carolina State University">South Carolina State University</option>
					<option value="South Central University of Econ. & Law">South Central University of Econ. & Law</option>
					<option value="South Dakota State University">South Dakota State University</option>
					<option value="South Suburban College">South Suburban College</option>
					<option value="Southeast Missouri State University">Southeast Missouri State University</option>
					<option value="Southeastern College Assemblies of God">Southeastern College Assemblies of God</option>
					<option value="Southern Adventist University">Southern Adventist University</option>
					<option value="Southern Connecticut State University">Southern Connecticut State University</option>
					<option value="Southern Illinois University Carbondale">Southern Illinois University Carbondale</option>
					<option value="Southern Illinois University Edwardsville">Southern Illinois University Edwardsville</option>
					<option value="Southern Methodist University">Southern Methodist University</option>
					<option value="Southern New Hampshire University">Southern New Hampshire University</option>
					<option value="Southern Oregon State University">Southern Oregon State University</option>
					<option value="Southern Oregon University">Southern Oregon University</option>
					<option value="Southern University at Baton Rouge">Southern University at Baton Rouge</option>
					<option value="Southern University Carbondale">Southern University Carbondale</option>
					<option value="Southern University of New Orleans">Southern University of New Orleans</option>
					<option value="Southwest Missouri State University">Southwest Missouri State University</option>
					<option value="Southwest Texas State University">Southwest Texas State University</option>
					<option value="Southwest University for Nationalities">Southwest University for Nationalities</option>
					<option value="Southwestern University Neofit Rilski">Southwestern University "Neofit Rilski"</option>
					<option value="Spelman College">Spelman College</option>
					<option value="Spertus College">Spertus College</option>
					<option value="Spoon River College">Spoon River College</option>
					<option value="Spring Arbor University">Spring Arbor University</option>
					<option value="Spring Arbor College">Spring Arbor College</option>
					<option value="Spring Hill College">Spring Hill College</option>
					<option value="Springfield College">Springfield College</option>
					<option value="SRM University">SRM University</option>
					<option value="Syracuse University">Syracuse University</option>
					<option value="St Ambrose University">St Ambrose University</option>
					<option value="St. Andrew (United Kingdom)">St. Andrew (United Kingdom)</option>
					<option value="St. Anthony College of Nursing">St. Anthony College of Nursing</option>
					<option value="St Augustine College">St Augustine College</option>
					<option value="St. Bonaventure University">St. Bonaventure University</option>
					<option value="St Johns University">St Johns University</option>
					<option value="St Joseph University">St Joseph University</option>
					<option value="St Louis University">St Louis University</option>
					<option value="St Marys College">St Marys College</option>
					<option value="St Marys College Notre Dame">St Marys College Notre Dame</option>
					<option value="St Marys University">St Marys University</option>
					<option value="St Norbert College">St Norbert College</option>
					<option value="St. Xavier College, Mumbai">St. Xavier College, Mumbai</option>
					<option value="St Xavier University">St Xavier University</option>
					<option value="St. Ambrose University">St. Ambrose University</option>
					<option value="St. Augustine College">St. Augustine College</option>
					<option value="St. Catherine University">St. Catherine University</option>
					<option value="St. Cloud State University">St. Cloud State University</option>
					<option value="St. Edwards University">St. Edwards University</option>
					<option value="St. Francis De Sales Seminary">St. Francis De Sales Seminary</option>
					<option value="St. Johns College">St. Johns College</option>
					<option value="St. Johns College-Ireland">St. Johns College-Ireland</option>
					<option value="St. Johns University">St. Johns University</option>
					<option value="St. Joseph College">St. Joseph College</option>
					<option value="St. Josephs University">St. Josephs University</option>
					<option value="St. Louis University">St. Louis University</option>
					<option value="St. Mary of the Woods College">St. Mary of the Woods College</option>
					<option value="St. Marys College">St. Marys College</option>
					<option value="St. Marys College - California">St. Marys College - California</option>
					<option value="St. Marys University">St. Marys University</option>
					<option value="St. Marys University of Minnesota">St. Marys University of Minnesota</option>
					<option value="St. Norbert College">St. Norbert College</option>
					<option value="St. Olaf College">St. Olaf College</option>
					<option value="St. Xavier University">St. Xavier University</option>
					<option value="Stanford University">Stanford University</option>
					<option value="State University of New York">State University of New York</option>
					<option value="State University of New York Albany">State University of New York Albany</option>
					<option value="State University of New York Purchase">State University of New York Purchase</option>
					<option value="State University of New York Binghampton">State University of New York Binghampton</option>
					<option value="State University of New York Brockport">State University of New York Brockport</option>
					<option value="State University of New York Fredonia">State University of New York Fredonia</option>
					<option value="State University of New York New Paltz">State University of New York New Paltz</option>
					<option value="State University of New York Oneonta">State University of New York Oneonta</option>
					<option value="State University of New York Oswego">State University of New York Oswego</option>
					<option value="State University of New York Old Westbury">State University of New York Old Westbury</option>
					<option value="State University of New York Plattsburg">State University of New York Plattsburg</option>
					<option value="State University of New York Geneseo">State University of New York Geneseo</option>
					<option value="State University of New York-Buffalo">State University of New York-Buffalo</option>
					
					<option value="State University of New York Stony Brook">State University of New York Stony Brook</option>
					<option value="Stephen Austin State University">Stephen Austin State University</option>
					<option value="Stephens College">Stephens College</option>
					<option value="Stetson University">Stetson University</option>
					<option value="Stonehill College">Stonehill College</option>
					<option value="Stony brook University">Stony brook University</option>
					<option value="Sungkyul Christian">Sungkyul Christian</option>
					<option value="SUNY College at Brockport">SUNY College at Brockport</option>
					<option value="Superior School Nueva">Superior School Nueva</option>
					<option value="Swarthmore College">Swarthmore College</option>
					
					<option value="Tabor College">Tabor College</option>
					<option value="Taiyuan University of Science & Tech">Taiyuan University of Science & Tech</option>
					<option value="Taylor University">Taylor University</option>
					<option value="Temple University">Temple University</option>
					<option value="Tennessee State University">Tennessee State University</option>
					<option value="Texas A & M University">Texas A & M University</option>
					<option value="Texas Christian University">Texas Christian University</option>
					<option value="Texas Southern University">Texas Southern University</option>
					<option value="Texas State University">Texas State University</option>
					<option value="Texas Tech University">Texas Tech University</option>
					<option value="Texas Womans Unversity">Texas Womans Unversity</option>
					<option value="The College of William and Mary">The College of William and Mary</option>
					<option value="The College of Wooster">The College of Wooster</option>
					<option value="The Polytechnic Ibadan">The Polytechnic Ibadan</option>
					<option value="Thomas Edison State College">Thomas Edison State College</option>
					<option value="Tianjin Normal University">Tianjin Normal University</option>
					<option value="Tiffin University">Tiffin University</option>
					<option value="TongJi University">TongJi University</option>
					<option value="Touro College">Touro College</option>
					<option value="Touro Lander for Women">Touro Lander for Women</option>
					<option value="Towson University">Towson University</option>
					<option value="Transylvania University">Transylvania University</option>
					<option value="Trent University">Trent University</option>
					<option value="Trinity Christian College">Trinity Christian College</option>
					<option value="Trinity College">Trinity College</option>
					<option value="Trinity College of Hartford, CT">Trinity College of Hartford, CT</option>
					<option value="Trinity International University">Trinity International University</option>
					<option value="Trinity Western University">Trinity Western University</option>
					<option value="Triton College">Triton College</option>
					<option value="Troy State University">Troy State University</option>
					<option value="Truman College">Truman College</option>
					<option value="Truman State">Truman State</option>
					<option value="Truman State University">Truman State University</option>
					<option value="Tufts University">Tufts University</option>
					<option value="Tulane University">Tulane University</option>
					<option value="Tung Hai University">Tung Hai University</option>
					<option value="Tuskegee University">Tuskegee University</option>
					<option value="Tyndale University">Tyndale University</option>
					
					<option value="U of C, Berkeley">U of C, Berkeley</option>
					<option value="U of Ghana">U of Ghana</option>
					
					<option value="U of KY">U of KY</option>
					<option value="U of MD">U of MD</option>
					<option value="U of Michigan">U of Michigan</option>
					<option value="U of Wisconsin Whitewater">U of Wisconsin Whitewater</option>
					<option value="U of Wisc-Parkside">U of Wisc-Parkside</option>
					<option value="U. of Mich.">U. of Mich.</option>
					<option value="U. of Port">U. of Port</option>
					<option value="U. of Wisc.- Eau Claire">U. of Wisc.- Eau Claire</option>
					<option value="UCLA">UCLA</option>
					
					
					
					<option value="Umass-Amherst">Umass-Amherst</option>
					<option value="Union College">Union College</option>
					<option value="Union University">Union University</option>
					<option value="United International College">United International College</option>
					<option value="University of Albany">University of Albany</option>
					<option value="University of Arkansas at Pine Bluff">University of Arkansas at Pine Bluff</option>
					<option value="University of Northern Iowa">University of Northern Iowa</option>
					<option value="Universidad Autonoma de SLP">Universidad Autonoma de SLP</option>
					<option value="Universidad de los Andes">Universidad de los Andes</option>
					<option value="Universidad Autonoma de Nuevo Leon">Universidad Autonoma de Nuevo Leon</option>
					
					<option value="University of Indianapolis">University of Indianapolis</option>
					<option value="University of Mississippi Main Campus">University of Mississippi Main Campus</option>
					<option value="University of Northern Iowa">University of Northern Iowa</option>
					<option value="University of South Carolina">University of South Carolina</option>
					<option value="University of Wisconsin Oshkosh">University of Wisconsin Oshkosh</option>
					<option value="University Ilorin">University Ilorin</option>
					<option value="University of Iowa">University of Iowa</option>
					<option value="University of Akron">University of Akron</option>
					<option value="University of Alabama">University of Alabama</option>
					<option value="University of Alabama Birmingham">University of Alabama Birmingham</option>
					<option value="University of Alabama Tuscaloosa">University of Alabama Tuscaloosa</option>
					<option value="University of Alaska">University of Alaska</option>
					<option value="University of Allahabad">University of Allahabad</option>
					<option value="University of Arizona">University of Arizona</option>
					<option value="University of Arkansas Fayettville">University of Arkansas Fayettville</option>
					<option value="University of Arkansas Pine Bluff">University of Arkansas Pine Bluff</option>
					<option value="University of Bamberg">University of Bamberg</option>
					<option value="University of Benin">University of Benin</option>
					<option value="University of Bombay">University of Bombay</option>
					<option value="University of Calabar">University of Calabar</option>
					<option value="University of California">University of California</option>
					<option value="University of California Irvine">University of California Irvine</option>
					<option value="University of California Los Angeles">University of California Los Angeles</option>
					<option value="University of California Berkeley">University of California Berkeley</option>
					<option value="University of California Davis">University of California Davis</option>
					<option value="University of California Irvine">University of California Irvine</option>
					<option value="University of California Riverside">University of California Riverside</option>
					<option value="University of California San Diego">University of California San Diego</option>
					<option value="University of California Santa Barbara">University of California Santa Barbara</option>
					<option value="University of California Santa Cruz">University of California Santa Cruz</option>
					<option value="University of Central Arkansas">University of Central Arkansas</option>
					<option value="University of Central Florida">University of Central Florida</option>
					<option value="University of Central Oklahoma">University of Central Oklahoma</option>  
					<option value="University of Central Missouri">University of Central Missouri</option>                   
					<option value="University of Chicago">University of Chicago</option>
					<option value="University of Cincinnati">University of Cincinnati</option>
					<option value="University of Colorado">University of Colorado</option>
					<option value="University of Colorado Boulder">University of Colorado Boulder</option>
					<option value="University of Connecticut">University of Connecticut</option>
					<option value="University of Dallas">University of Dallas</option>
					<option value="University of Dar Es Salaam">University of Dar Es Salaam</option>
					<option value="University of Dayton">University of Dayton</option>
					<option value="University of Delaware">University of Delaware</option>
					<option value="University of Delhi">University of Delhi</option>
					<option value="University of Denver">University of Denver</option>
					<option value="University of Detroit Mercy">University of Detroit Mercy</option>
					<option value="University of Development Studies">University of Development Studies</option>
					<option value="University of Dhaka">University of Dhaka</option>
					<option value="University of Evansville">University of Evansville</option>
					<option value="University of Findlay">University of Findlay</option>
					<option value="University of Florida">University of Florida</option>
					<option value="University of Florida Gainesville">University of Florida Gainesville</option>
					<option value="University of Georgia">University of Georgia</option>
					<option value="University of Georgia-Athens">University of Georgia-Athens</option>
					<option value="University of Ghana">University of Ghana</option>
					<option value="University of Haifa">University of Haifa</option>
					<option value="University of Hartford">University of Hartford</option>
					<option value="University of Hawaii">University of Hawaii</option>
					<option value="University of Hawaii Manoa">University of Hawaii Manoa</option>
					<option value="University of Hong Kong">University of Hong Kong</option>
					<option value="University of Houston">University of Houston</option>
					<option value="University of Houston-Downtown">University of Houston-Downtown</option>
					<option value="University of Ibadan">University of Ibadan</option>
					<option value="University of Illinois Urbana">University of Illinois Urbana</option>
					<option value="University of Illinois Chicago">University of Illinois Chicago</option>
					<option value="University of Illinois Springfield">University of Illinois Springfield</option>
					
					<option value="University of Illinois Urbana Champaign">University of Illinois Urbana Champaign</option>
					<option value="University of Incarnate Word">University of Incarnate Word</option>
					<option value="University of Indiana">University of Indiana</option>
					<option value="University of Indianapolis">University of Indianapolis</option>
					<option value="University of Indiana-South Bend">University of Indiana-South Bend</option>
					<option value="University of Iowa">University of Iowa</option>
					<option value="University of Iowa College of Liberal Arts and Sci">University of Iowa College of Liberal Arts and Sci</option>
					<option value="University of Jamestown">University of Jamestown</option>
					<option value="University of Kansas">University of Kansas</option>
					<option value="University of Kansas Lawrence">University of Kansas Lawrence</option>
					<option value="University of Karachi">University of Karachi</option>
					<option value="University of Kentucky">University of Kentucky</option>
					<option value="University of Liberia">University of Liberia</option>
					<option value="University of Louisiana Lafayette">University of Louisiana Lafayette</option>
					<option value="University of Louisville">University of Louisville</option>
					<option value="University of Lucknow">University of Lucknow</option>
					<option value="University of Maine">University of Maine</option>
				   
					<option value="University of Mary Washington">University of Mary Washington</option>
					<option value="University of Maryland">University of Maryland</option>
					<option value="University of Maryland-Baltimore County">University of Maryland-Baltimore County</option>
					<option value="University of Maryland College Park">University of Maryland College Park</option>
					<option value="University of Massachusettes Lowell">University of Massachusettes Lowell</option>
					<option value="university of Massachusetts">university of Massachusetts</option>
					<option value="University of Massachusetts Amherst">University of Massachusetts Amherst</option>
					<option value="University of Massachusetts Boston">University of Massachusetts Boston</option>
					<option value="University of Melbourne">University of Melbourne</option>
					<option value="University of Memphis">University of Memphis</option>
					<option value="University of Miami">University of Miami</option>
			  
					<option value="University of Michigan">University of Michigan</option>
					<option value="University of Michigan Ann Arbor">University of Michigan Ann Arbor</option>
					<option value="University of Michigan Flint">University of Michigan Flint</option>
					<option value="University of Michigan-Dearborn">University of Michigan-Dearborn</option>
					<option value="University of Minnesota">University of Minnesota</option>
					<option value="University of Minnesota Twin Cities">University of Minnesota Twin Cities</option>
					<option value="University of Minnesota, Duluth">University of Minnesota, Duluth</option>
					<option value="University of Mississippi">University of Mississippi</option>
					<option value="University of Missouri">University of Missouri</option>
					<option value="University of Missouri - St. Louis">University of Missouri - St. Louis</option>
					<option value="University of Missouri Columbia">University of Missouri Columbia</option>
					
					<option value="University of Missouri Kansas City">University of Missouri Kansas City</option>
					<option value="University of Montana">University of Montana</option>
					<option value="University of Nebraska">University of Nebraska</option>
					<option value="University of Nebraska Kearney">University of Nebraska Kearney</option>
					<option value="University of Nebraska Lincoln">University of Nebraska Lincoln</option>
					<option value="University of Nebraska Omaha">University of Nebraska Omaha</option>
					
					<option value="University of Nevada">University of Nevada</option>
					<option value="University of Nevada Las Vegas">University of Nevada Las Vegas</option>
					<option value="University of Nevada Reno">University of Nevada Reno</option>
					<option value="University of New England">University of New England</option>
					<option value="University of New Hampshire">University of New Hampshire</option>
					<option value="University of New Hampshire Manchester">University of New Hampshire Manchester</option>
					<option value="University of New Mexico">University of New Mexico</option>
					<option value="University of New Mexico Albuquerque">University of New Mexico Albuquerque</option>
					<option value="University of New Orleans">University of New Orleans</option>
					<option value="University of Nigeria">University of Nigeria</option>
					<option value="University of Norte Dame">University of Norte Dame</option>
					<option value="University of North Alabama">University of North Alabama</option>
					<option value="University of North Carolina">University of North Carolina</option>
					<option value="University of North Carolina Asheville">University of North Carolina Asheville</option>
					<option value="University of North Carolina Chapel Hill">University of North Carolina Chapel Hill</option>
					<option value="University of North Carolina Charlotte">University of North Carolina Charlotte</option>
					<option value="University of North Carolina Greensboro">University of North Carolina Greensboro</option>
					<option value="University of North Carolina Pembroke">University of North Carolina Pembroke</option>
					<option value="University of North Carolina Wilmington">University of North Carolina Wilmington</option>
					<option value="University of North Dakota">University of North Dakota</option>
					<option value="University of North Texas">University of North Texas</option>
					<option value="university of Northern Colorado">university of Northern Colorado</option>
					<option value="University of Northern Iowa">University of Northern Iowa</option>
					<option value="University of Notre Dame">University of Notre Dame</option>
					<option value="University of Oklahoma">University of Oklahoma</option>
					<option value="University of Oklahoma Norman">University of Oklahoma Norman</option>
					<option value="University of Oregon">University of Oregon</option>
					<option value="University of Oregon Main Campus">University of Oregon Main Campus</option>
					<option value="University of Pennsylvania">University of Pennsylvania</option>
					<option value="University of Phoenix">University of Phoenix</option>
					<option value="University of Pitesti">University of Pitesti</option>
					<option value="University of Pittsburgh">University of Pittsburgh</option>
					<option value="University of Pittsburgh Main Campus">University of Pittsburgh Main Campus</option>
					<option value="University of Portland">University of Portland</option>
					<option value="University of Puerto Rico">University of Puerto Rico</option>
					<option value="University of Puget Sound">University of Puget Sound</option>
					<option value="University of Pune">University of Pune</option>
					<option value="University of Redlands">University of Redlands</option>
					<option value="University of Rhode Island">University of Rhode Island</option>
					<option value="University of Richmond">University of Richmond</option>
					<option value="University of Rochester">University of Rochester</option>
					<option value="University of Saint Thomas">University of Saint Thomas</option>
					<option value="University of Salsbury">University of Salsbury</option>
					<option value="University of San Diego">University of San Diego</option>
					<option value="University of San Francisco">University of San Francisco</option>
					<option value="University of San Paulo">University of San Paulo</option>
					<option value="University of Scranton">University of Scranton</option>
					<option value="University of Sioux Falls">University of Sioux Falls</option>
					<option value="University of South Carolina">University of South Carolina</option>
					<option value="University of South Carolina Columbia">University of South Carolina Columbia</option>
					<option value="University of South Dakota">University of South Dakota</option>
					<option value="University of South Florida">University of South Florida</option>
					<option value="University of Southern California">University of Southern California</option>
					<option value="University of Southern Florida">University of Southern Florida</option>
					<option value="University of Southern Indiana">University of Southern Indiana</option>
					<option value="University of Southern Mississippi">University of Southern Mississippi</option>
					<option value="University of Southwestern Louisana">University of Southwestern Louisana</option>
					<option value="University of St. Andrews (Scotland)">University of St. Andrews (Scotland)</option>
					<option value="University of St Thomas">University of St Thomas</option>
					<option value="University of St. Francis">University of St. Francis</option>
					<option value="University of St. Francis - Indiana">University of St. Francis - Indiana</option>
					<option value="University of St. Iowa">University of St. Iowa</option>
					<option value="University of St. Maine">University of St. Maine</option>
					<option value="University of St. Michigan">University of St. Michigan</option>
					<option value="University of St. Missouri St. Louis">University of St. Missouri St. Louis</option>
					<option value="University of St. Thomas">University of St. Thomas</option>
					<option value="University of St. Wisconsin-Madison">University of St. Wisconsin-Madison</option>
					<option value="University of Tampa">University of Tampa</option>
					<option value="University of Tennessee">University of Tennessee</option>
					<option value="University of Tennessee Knoxville">University of Tennessee Knoxville</option>
					<option value="University of Texas">University of Texas</option>
					<option value="University of Texas Austin">University of Texas Austin</option>
					<option value="University of Texas Arlington">University of Texas Arlington</option>
					<option value="University of Texas Dallas">University of Texas Dallas</option>
					<option value="University of Texas El Paso">University of Texas El Paso</option>
					<option value="University of Texas Pan American">University of Texas Pan American</option>
					<option value="University of the Ozarks">University of the Ozarks</option>
					<option value="University of the West Indies">University of the West Indies</option>
					<option value="University of the Witwatersraud">University of the Witwatersraud</option>
					<option value="University of Tibet">University of Tibet</option>
					<option value="University of Toledo">University of Toledo</option>
					<option value="University of Tulsa">University of Tulsa</option>
					<option value="University of Utah">University of Utah</option>
					<option value="University of Vermont">University of Vermont</option>
					<option value="university of Virginia">university of Virginia</option>
					<option value="University of Washington">University of Washington</option>
					<option value="University of Washington Seattle">University of Washington Seattle</option>
					<option value="University of West Florida">University of West Florida</option>
					<option value="University of West Indies">University of West Indies</option>
					<option value="University of Western Ontario">University of Western Ontario</option>
					<option value="University of Windsor">University of Windsor</option>
					<option value="University of Windsor Canada">University of Windsor Canada</option>
					<option value="University of Winnipeg">University of Winnipeg</option>
					<option value="University of Wisconsin">University of Wisconsin</option>
					<option value="University of Wisconsin Madison">University of Wisconsin Madison</option>
					<option value="University of Wisconsin Oshkosh">University of Wisconsin Oshkosh</option>
					<option value="University of Wisconsin Eau Claire">University of Wisconsin Eau Claire</option>
					<option value="University of Wisconsin Green Bay">University of Wisconsin Green Bay</option>
					<option value="University of Wisconsin La Crosse">University of Wisconsin La Crosse</option>
					<option value="University of Wisconsin LaCrosse">University of Wisconsin LaCrosse</option>
					<option value="University of Wisconsin Milwaukee">University of Wisconsin Milwaukee</option>
					<option value="University of Wisconsin Parkside">University of Wisconsin Parkside</option>
					<option value="University of Wisconsin Platteville">University of Wisconsin Platteville</option>
					
					<option value="University of Wisconsin Stevens Point">University of Wisconsin Stevens Point</option>
					<option value="University of Wisconsin Stout">University of Wisconsin Stout</option>
					<option value="University of Wisconsin Whitewater">University of Wisconsin Whitewater</option>
					<option value="University of Wyoming">University of Wyoming</option>
					<option value="University of Zululand">University of Zululand</option>
					<option value="Univertsity of St. Francis">Univertsity of St. Francis</option>
					<option value="Uniwersytet Adama">Uniwersytet Adama</option>
					<option value="Upper Iowa University">Upper Iowa University</option>
					<option value="Utah State University">Utah State University</option>
					<option value="Utah Valley State College">Utah Valley State College</option>
					<option value="Utah Valley University">Utah Valley University</option>
					<option value="Valdosta State University">Valdosta State University</option>
					<option value="Valparaiso University">Valparaiso University</option>
					<option value="Vanderbilt University">Vanderbilt University</option>
					<option value="Vanguard University of Southern California">Vanguard University of Southern California</option>
					<option value="Vassar College">Vassar College</option>
					<option value="Villanova University">Villanova University</option>
					<option value="Virginia Commonwealth University">Virginia Commonwealth University</option>
					<option value="Virginia Polytech Institute">Virginia Polytech Institute</option>
					<option value="Viterbo University">Viterbo University</option>
					<option value="Wake Forest University">Wake Forest University</option>
					<option value="Warren Wilson College">Warren Wilson College</option>
					<option value="Wartburg College">Wartburg College</option>
					<option value="Washburn University">Washburn University</option>
					<option value="Washington State University">Washington State University</option>
					<option value="Washington University St. Louis">Washington University St. Louis</option>
					<option value="Waubonsee Community College">Waubonsee Community College</option>
					<option value="Wayne State University">Wayne State University</option>
					<option value="Webster University">Webster University</option>
					<option value="Wells College">Wells College</option>
					<option value="Wesleyan University">Wesleyan University</option>
					<option value="West Chester University">West Chester University</option>
					<option value="West Chester University of Pennsylvania">West Chester University of Pennsylvania</option>
					<option value="West Virginia University">West Virginia University</option>
					<option value="Western Carolina University">Western Carolina University</option>
					<option value="Western Connecticut State University">Western Connecticut State University</option>
					<option value="Western Illinois University">Western Illinois University</option>
					<option value="Western Kentucky University">Western Kentucky University</option>
					<option value="Western Michigan State University">Western Michigan State University</option>
					<option value="Western Michigan University">Western Michigan University</option>
					<option value="Western Washington University">Western Washington University</option>
					<option value="Westminster College">Westminster College</option>
					<option value="Westmont College">Westmont College</option>
					<option value="Westwood College">Westwood College</option>
					<option value="Wheaton College">Wheaton College</option>
					<option value="Wheelock College">Wheelock College</option>
					<option value="Winston-Salem State University">Winston-Salem State University</option>
					<option value="Whitman College">Whitman College</option>
					
					<option value="Whittier College">Whittier College</option>
					<option value="Whitworth University">Whitworth University</option>
					<option value="Wichita State University">Wichita State University</option>
					<option value="Wilbur Wright College">Wilbur Wright College</option>
					<option value="Willam Woods University">Willam Woods University</option>
					<option value="Willamette University">Willamette University</option>
					<option value="William Penn University">William Penn University</option>
					<option value="William Smith College">William Smith College</option>
					<option value="William Woods University">William Woods University</option>
					<option value="Winona State University">Winona State University</option>
					<option value="Winthrop University">Winthrop University</option>
					<option value="Wittenberg University">Wittenberg University</option>
					<option value="Wright State University">Wright State University</option>
					<option value="Wuhan University">Wuhan University</option>
					<option value="Xavier University">Xavier University</option>
					<option value="Xavier University Louisiana">Xavier University Louisiana</option>
					<option value="Yale University">Yale University</option>
					<option value="Yancheng Teachers University">Yancheng Teachers University</option>
					<option value="Yeshiva University">Yeshiva University</option>
					<option value="Yonsei University">Yonsei University</option>
					<option value="York College">York College</option>
					<option value="York College of Pennsylvania">York College of Pennsylvania</option>
					<option value="York University">York University</option>
					<option value="Zhejiang University">Zhejiang University</option>
					</select>
					<br/><br/><br/><br/>
					<label>UG GPA</label>
					<input type="text" name="uggpa" class="gpa" id="uggpa" value='<%Response.write rs("UGGPA") %>' />
					<label>UG Major</label>
					<input type="text" name="ugmajor" id="ugmajor" value="<%Response.write rs("UGMajor") %>" />                    
					<label>Grad College</label>
					<select name="gradcollege" id="gradcollege">
					<option value="<%= rs.Fields(55) %>"><%= rs.Fields(55) %></option>    
					<option value="Adam Mickiewicz University">Adam Mickiewicz University</option>
					<option value="Addis Ababa">Addis Ababa</option>
					<option value="American Miltary Unversity">American Miltary Unversity</option>
					<option value="Andrews University">Andrews University</option>
					<option value="Ann Blitstein Institute">Ann Blitstein Institute</option>
					<option value="Argosy University">Argosy University</option>
					<option value="Arizona State University">Arizona State University</option>
					<option value="Atlanta University">Atlanta University</option>
					<option value="Aurora University">Aurora University</option>
					<option value="Bank St College of Educ">Bank St College of Educ</option>
					<option value="Beijing Institute of Fashion Technology">Beijing Institute of Fashion Technology</option>
					<option value="Bellevue University">Bellevue University</option>
					<option value="Bhavnagar University">Bhavnagar University</option>
					<option value="Boston College">Boston College</option>
					<option value="Boston University">Boston University</option>
					<option value="Bowdoin College">Bowdoin College</option>
					<option value="Bradley University">Bradley University</option>   
					<option value="British American University">British American University</option>
					<option value="California State University">California State University</option>
					<option value="California State University-Northridge">California State University-Northridge</option>
					<option value="Capitol Normal University, Beijing">Capitol Normal University, Beijing</option>
					<option value="Carlton University">Carlton University</option>
					<option value="Case Western Reserve University">Case Western Reserve University</option>
					<option value="Catholic Theological Union">Catholic Theological Union</option>
					<option value="Catholic University of America">Catholic University of America</option>
					<option value="Chapman University">Chapman University</option>
					<option value="Charles University of Prague">Charles University of Prague</option>
					<option value="Chicago Kent College of Law">Chicago Kent College of Law</option>
					<option value="Chicago State University">Chicago State University</option>
					<option value="ChonbukNational University">ChonbukNational University</option>
					<option value="Chung-Ang University">Chung-Ang University</option>
					<option value="Clark Atlanta University">Clark Atlanta University</option>
					<option value="Colorado State University">Colorado State University</option>
					<option value="Columbia University">Columbia University</option>
					<option value="Concordia University">Concordia University</option>
					<option value="CSU-Long Beach">CSU-Long Beach</option>
					<option value="CUNY school of Law">CUNY school of Law</option>
					<option value="Dankook University">Dankook University</option>
					<option value="De Paul University">De Paul University</option>
					<option value="Devi Ahilya University Idore (Madhya Pradesh, India)">Devi Ahilya University Idore (Madhya Pradesh, India)</option>
					<option value="Devry University">Devry University</option>
					<option value="Dominican University">Dominican University</option>
					<option value="Dowling College">Dowling College</option>
					<option value="Drake University">Drake University</option>
					<option value="Duke Divinity School">Duke Divinity School</option>
					<option value="East China Normal University">East China Normal University</option>
					<option value="Eastern Illinois University">Eastern Illinois University</option>
					<option value="Emmanuel College">Emmanuel College</option>
					<option value="Erickson Institue">Erickson Institue</option>
					<option value="EWAH Womens University">EWAH Womens University</option>
					<option value="Florida International University">Florida International University</option>
					<option value="Fordham University">Fordham University</option>
					<option value="Fuller Theological Seminary">Fuller Theological Seminary</option>
					<option value="Fullerton College">Fullerton College</option>
					<option value="Garrett Evangelical Seminary">Garrett Evangelical Seminary</option>
					<option value="George Mason University">George Mason University</option>
					<option value="George Washington Univ.">George Washington Univ.</option>
					<option value="George Williams College">George Williams College</option>
					<option value="George Williams University">George Williams University</option>
					<option value="Georgetown University">Georgetown University</option>
					<option value="Governors State University">Governors State University</option>
					<option value="Grand Valley State University">Grand Valley State University</option>
					<option value="Harold Washington College">Harold Washington College</option>
					<option value="Hebrew University">Hebrew University</option>
					<option value="Hunter College">Hunter College</option>
					<option value="Illinois State University">Illinois State University</option>
					<option value="Indiana State University">Indiana State University</option>
					<option value="Indiana University">Indiana University</option>
					<option value="Indiana University of Pennsylvania">Indiana University of Pennsylvania</option>
					<option value="Istanbul Univrsity">Istanbul Univrsity</option>
					<option value="ISWR-Dacca University-Bangladesh">ISWR-Dacca University-Bangladesh</option>
					<option value="Jane Addams College of Social Work">Jane Addams College of Social Work</option>
					<option value="Jawaharlal Nehru University">Jawaharlal Nehru University</option>
					<option value="John Marshall Law School">John Marshall Law School</option>
					<option value="Johns Hopkins University">Johns Hopkins University</option>
					<option value="Kansai Gaidai University">Kansai Gaidai University</option>
					<option value="Keller Graduate School of Management">Keller Graduate School of Management</option>
					<option value="Kennesaw State University">Kennesaw State University</option>
					<option value="Kent College of Law">Kent College of Law</option>
					<option value="Kentucky Wesleyan College">Kentucky Wesleyan College</option>
					<option value="Lewis University">Lewis University</option>
					<option value="Liberty University">Liberty University</option>
					<option value="Lindenwood University-Belleville">Lindenwood University-Belleville</option>
					<option value="Louisana State University">Louisana State University</option>
					<option value="Loyola University Chicago">Loyola University Chicago</option>
					<option value="Loyola University Maryland">Loyola University Maryland</option>
					<option value="Loyola University of Maryland">Loyola University of Maryland</option>
					<option value="Mahatma Gandhi University">Mahatma Gandhi University</option>
					<option value="Marshall University">Marshall University</option>
					<option value="Marywood University">Marywood University</option>
					<option value="McCormick Theological Seminary">McCormick Theological Seminary</option>
					<option value="Metropolitan State College of Denver">Metropolitan State College of Denver</option>
					<option value="Miami University">Miami University</option>
					<option value="Miami University-Oxford Ohio">Miami University-Oxford Ohio</option>
					<option value="Miami of Ohio University">Miami of Ohio University</option>
					<option value="Michigan State University">Michigan State University</option>
					<option value="Midwest Christian College">Midwest Christian College</option>
					<option value="Milligan College">Milligan College</option>
					<option value="Moody Bible Institute">Moody Bible Institute</option>
					<option value="National Louis University">National Louis University</option>
					<option value="Nazarene Theological">Nazarene Theological</option>
					<option value="New Mexico Highlands University">New Mexico Highlands University</option>
					<option value="New York University">New York University</option>
					<option value="New York University School of Law">New York University School of Law</option>
					<option value="Norfolk State University">Norfolk State University</option>
					<option value="North Park University">North Park University</option>
					<option value="Northern Arizona University">Northern Arizona University</option>
					<option value="Northeastern Illinois University">Northeastern Illinois University</option>
					<option value="Northern Illinois University">Northern Illinois University</option>
					<option value="Northwestern University">Northwestern University</option>
					<option value="Northwestern College">Northwestern College</option>
					<option value="Nottingham Trent University">Nottingham Trent University</option>
					<option value="Oakton Community College">Oakton Community College</option>
					<option value="Olivet Nazarene University">Olivet Nazarene University</option>
					<option value="Parsons The New School for Design">Parsons The New School for Design</option>
					<option value="Penn State Univ">Penn State Univ</option>
					<option value="Pomona College">Pomona College</option>
					<option value="Providence College">Providence College</option>
					<option value="PuKyong National University">PuKyong National University</option>
					<option value="Qingdao Binhai University">Qingdao Binhai University</option>
					<option value="Reading University">Reading University</option>
					<option value="Regent University">Regent University</option>
					<option value="Regis University">Regis University</option>
					<option value="Roberts Wesleyan College">Roberts Wesleyan College</option>
					<option value="Roosevelt University">Roosevelt University</option>
					<option value="Rutgers University">Rutgers University</option>
					<option value="San Francisco State">San Francisco State</option>
					<option value="School of the Art Institute of Chicago">School of the Art Institute of Chicago</option>
					<option value="Sewanee: The University of The South">Sewanee: The University of The South</option>
					<option value="Smith College">Smith College</option>
					<option value="Smith College School of Social Work">Smith College School of Social Work</option>
					<option value="Soongsil University">Soongsil University</option>
					<option value="Southern Baptist Theological Seminary">Southern Baptist Theological Seminary</option>
					<option value="Southern Illinois University">Southern Illinois University</option>
					<option value="Southern Methodist University">Southern Methodist University</option>
					<option value="Southern New Hampshire University">Southern New Hampshire University</option>
					<option value="Southwestern University Neofit Rilski">Southwestern University "Neofit Rilski"</option>
					<option value="Spertus College">Spertus College</option>
					<option value="Springfield College">Springfield College</option>
					<option value="St. Andrew (United Kingdom)">St. Andrew (United Kingdom)</option>
					<option value="St. Augustine College">St. Augustine College</option>
					<option value="St. Bonaventure University">St. Bonaventure University</option>
					<option value="St. Leo University">St. Leo University</option>
					<option value="St. Louis University">St. Louis University</option>
					<option value="St. Xavier College, Mumbai">St. Xavier College, Mumbai</option>
					<option value="St. Xavier University">St. Xavier University</option>
					<option value="SUNY College at Brockport">SUNY College at Brockport</option>
					<option value="Tata Institute">Tata Institute</option>
					<option value="Tata Institute India">Tata Institute India</option>
					<option value="Tata Institute of Social Sciences">Tata Institute of Social Sciences</option>
					<option value="Taylor University">Taylor University</option>
					<option value="Tel Aviv University">Tel Aviv University</option>
					<option value="Temple University">Temple University</option>
					<option value="The Catholic University of America">The Catholic University of America</option>
					<option value="The College of William and Mary">The College of William and Mary</option>
					<option value="The George Washington University">The George Washington University</option>
					<option value="The University of Jerusalem">The University of Jerusalem</option>
					<option value="Thomas Cooley Law School">Thomas Cooley Law School</option>
					<option value="Troy State University">Troy State University</option>
					<option value="Tulane University">Tulane University</option>
					<option value="Tyndale University">Tyndale University</option>
					<option value="Union Theological Seminary">Union Theological Seminary</option>
					<option value="Univeristy of Guadalajara">Univeristy of Guadalajara</option>
					<option value="Universidad Pontificia de Comillas">Universidad Pontificia de Comillas</option>
					<option value="University for Peace - Costa Rica">University for Peace - Costa Rica</option>
					<option value="University of Ado-Ekiti">University of Ado-Ekiti</option>
					<option value="University of Akron">University of Akron</option>
					<option value="University of Alabama">University of Alabama</option>
					<option value="University of Alabama Birmingham">University of Alabama Birmingham</option>
					<option value="University of Albany">University of Albany</option>
					<option value="University of Arkansas at Pine Bluff">University of Arkansas at Pine Bluff</option>
					<option value="University of Benin">University of Benin</option>
					<option value="University of Bombay-India">University of Bombay-India</option>
					<option value="University of Bristol">University of Bristol</option>
					<option value="University of California Berkley">University of California Berkley</option>
					<option value="University of California Los Angeles">University of California Los Angeles</option>
					<option value="University of California-Berkeley">University of California-Berkeley</option>
					<option value="University of California-Long Beach">University of California-Long Beach</option>
					<option value="University of Central Florida">University of Central Florida</option>
					<option value="University of Chicago">University of Chicago</option>
					<option value="University of Cincinnati">University of Cincinnati</option>
					<option value="University of Georgia">University of Georgia</option>
					<option value="University of Glasgow">University of Glasgow</option>
					<option value="University of Hawaii">University of Hawaii</option>
					<option value="University of Hong Kong">University of Hong Kong</option>
					<option value="University of Houston">University of Houston</option>
					<option value="University of Ibadan">University of Ibadan</option>
					<option value="University of Illinois">University of Illinois</option>
					<option value="University of Illinois Chicago">University of Illinois Chicago</option>
					<option value="University of Illinois Springfield">University of Illinois Springfield</option>
					<option value="University of Illinois Urbana Champaign">University of Illinois Urbana Champaign</option>
					<option value="University of Iowa">University of Iowa</option>
					<option value="University of Karachi">University of Karachi</option>
					<option value="University of Kentucky">University of Kentucky</option>
					<option value="University of Madras-India">University of Madras-India</option>
					<option value="University of Maryland">University of Maryland</option>
					<option value="University of Maryland-Baltimore County">University of Maryland-Baltimore County</option>
					<option value="University of Massachusetts Amherst">University of Massachusetts Amherst</option>
					<option value="University of Michigan">University of Michigan</option>
					<option value="University of Michigan Ann Arbor">University of Michigan Ann Arbor</option>
					<option value="University of Minnesota">University of Minnesota</option>
					<option value="University of Missouri Columbia">University of Missouri Columbia</option>
					<option value="University of Natal">University of Natal</option>
					<option value="University of North Carolina">University of North Carolina</option>
					<option value="University of North Carolina Chapel Hill">University of North Carolina Chapel Hill</option>
					<option value="University of North Carolina Charlotte">University of North Carolina Charlotte</option>
					<option value="University of North Carolina Wilmington">University of North Carolina Wilmington</option>
					<option value="University of Northern Iowa">University of Northern Iowa</option>
					<option value="University of Oregon">University of Oregon</option>
					<option value="University of Pennsylvania">University of Pennsylvania</option>
					<option value="University of Pittsburgh">University of Pittsburgh</option>
					<option value="University of South Carolina">University of South Carolina</option>
					<option value="University of South Florida">University of South Florida</option>
					<option value="University of Southern California">University of Southern California</option>
					<option value="University of Southern Maine">University of Southern Maine</option>
					<option value="University of St. Andrews (Scotland)">University of St. Andrews (Scotland)</option>
					<option value="University of Tennessee">University of Tennessee</option>
					<option value="University of Texas El Paso">University of Texas El Paso</option>
					<option value="University of Toledo">University of Toledo</option>
					<option value="University of Vermont">University of Vermont</option>
					<option value="University of Washington">University of Washington</option>
					<option value="University of Wisconsin">University of Wisconsin</option>
					<option value="University of Wisconsin Madison">University of Wisconsin Madison</option>
					<option value="University of Wisconsin Milwaukee">University of Wisconsin Milwaukee</option>
					<option value="Virginia Commonwealth University">Virginia Commonwealth University</option>
					<option value="Washington University">Washington University</option>
					<option value="Wayne State University">Wayne State University</option>
					<option value="Webster University">Webster University</option>
					<option value="Western Illinois University">Western Illinois University</option>
					<option value="Western Maryland College">Western Maryland College</option>
					<option value="Western Michigan University">Western Michigan University</option>
					<option value="Wheaton College">Wheaton College</option>
					<option value="Winston-Salem State University">Winston-Salem State University</option>
					<option value="Yale Divinity School">Yale Divinity School</option>
					<option value="Yale University">Yale University</option>
					<option value="Yeshiva University">Yeshiva University</option>
					<option value="Yonsie University">Yonsie University</option>
					<option value="York University">York University</option>
					</select>
					<br/><br/><br/><br/>
					<label>Grad GPA</label>
					<input type="text" name="gradgpa" class="gpa" id="gradgpa" value='<%Response.write rs("GradGPA") %>' />
					<label>Grad Major</label>
					<input type="text" name="gradmajor" id="gradmajor" value="<%Response.write rs("GradMajor") %>" />
					<label>Grad Degree</label>
					<input type="text" name="graddegree" id="graddegree" value='<%Response.write rs("GradDegree") %>' />
					
					<br/><br/><br/><br/>
					</p>
					<button type="submit" name="Submit" onclick="this.form.action='AfterEditStudent.asp?UIN=' + this.value; this.forms.submit();" value='<% Response.write rs("UIN") %>'>Save</button><br /><br />
					</fieldset>
					

					<div id="scholarship" align="center">
					<fieldset class="step">
					<legend></legend>
					<p>
					<br/>

					<label>Scholarship Type 1</label>
					<select name="award_type1" id="award_type1" >
					<option value="<%= rs.Fields(63) %>"><%= rs.Fields(63) %></option>
					<option value="Nancy J. Alexander Scholarship">Nancy J. Alexander Scholarship</option>
					<option value="Sandra and Furio Alberti Memorial Award">Sandra and Furio Alberti Memorial Award</option>
					<option value="Nancy Ali Memorial Scholarship">Nancy Ali Memorial Scholarship</option>
					<option value="Chicago Foundlings Home Award">Chicago Foundlings Home Award</option>
					<option value="Patricia Counce Memorial Scholarship/Delta Sigma Theta Sorority">Patricia Counce Memorial Scholarship/Delta Sigma Theta Sorority</option>
					<option value="Smith Barney Mercantile Foundation Award">Smith Barney Mercantile Foundation Award</option>
					<option value="Kott Award">Kott Award</option>
					<option value="Gabe Miller Memorial Foundation">Gabe Miller Memorial Foundation</option>
					<option value="Fall and Spring Board of Trustees Tuition and Fee Waivers">Fall and Spring Board of Trustees Tuition and Fee Waivers</option>
					<option value="Summer Board of Trustee Waivers">Summer Board of Trustee Waivers</option>
					<option value="Harvey Trejer Award">Harvey Trejer Award</option>
					<option value="Deans Scholarship">Deans Scholarship</option>
					</select>   
					<label>Award Amount</label>
					<input type="text" name="award_amount1" id="award_amount1" value='<%Response.write rs("award_amount") %>'/>
					
					<label>Award Date</label>
					<input type="text" name="award_date1" id="award_date1" class="date" value='<%Response.write rs("award_date") %>'/>
					<br/><br/><br/><br/>

					<label>Scholarship Type 2</label>
					<select name="award_type2" id="award_type2" >
					<option value="<%= rs.Fields(75) %>"><%= rs.Fields(75) %></option>
					<option value="Nancy J. Alexander Scholarship">Nancy J. Alexander Scholarship</option>
					<option value="Sandra and Furio Alberti Memorial Award">Sandra and Furio Alberti Memorial Award</option>
					<option value="Nancy Ali Memorial Scholarship">Nancy Ali Memorial Scholarship</option>
					<option value="Chicago Foundlings Home Award">Chicago Foundlings Home Award</option>
					<option value="Patricia Counce Memorial Scholarship/Delta Sigma Theta Sorority">Patricia Counce Memorial Scholarship/Delta Sigma Theta Sorority</option>
					<option value="Smith Barney Mercantile Foundation Award">Smith Barney Mercantile Foundation Award</option>
					<option value="Kott Award">Kott Award</option>
					<option value="Gabe Miller Memorial Foundation">Gabe Miller Memorial Foundation</option>
					<option value="Fall and Spring Board of Trustees Tuition and Fee Waivers">Fall and Spring Board of Trustees Tuition and Fee Waivers</option>
					<option value="Summer Board of Trustee Waivers">Summer Board of Trustee Waivers</option>
					<option value="Harvey Trejer Award">Harvey Trejer Award</option>
					<option value="Deans Scholarship">Deans Scholarship</option>
					</select>   
					<label>Award Amount</label>
					<input type="text" name="award_amount2" id="award_amount2" value='<%Response.write rs("award_amount2") %>'/>
					
					<label>Award Date</label>
					<input type="text" name="award_date2" id="award_date2" class="date" value='<%Response.write rs("award_date2") %>'/>
					<br/><br/><br/><br/>

					<label>Scholarship Type 3</label>
					<select name="award_type3" id="award_type3" >
					<option value="<%= rs.Fields(78) %>"><%= rs.Fields(78) %></option>
					<option value="Nancy J. Alexander Scholarship">Nancy J. Alexander Scholarship</option>
					<option value="Sandra and Furio Alberti Memorial Award">Sandra and Furio Alberti Memorial Award</option>
					<option value="Nancy Ali Memorial Scholarship">Nancy Ali Memorial Scholarship</option>
					<option value="Chicago Foundlings Home Award">Chicago Foundlings Home Award</option>
					<option value="Patricia Counce Memorial Scholarship/Delta Sigma Theta Sorority">Patricia Counce Memorial Scholarship/Delta Sigma Theta Sorority</option>
					<option value="Smith Barney Mercantile Foundation Award">Smith Barney Mercantile Foundation Award</option>
					<option value="Kott Award">Kott Award</option>
					<option value="Gabe Miller Memorial Foundation">Gabe Miller Memorial Foundation</option>
					<option value="Fall and Spring Board of Trustees Tuition and Fee Waivers">Fall and Spring Board of Trustees Tuition and Fee Waivers</option>
					<option value="Summer Board of Trustee Waivers">Summer Board of Trustee Waivers</option>
					<option value="Harvey Trejer Award">Harvey Trejer Award</option>
					<option value="Deans Scholarship">Deans Scholarship</option>
					</select>   
					<label>Award Amount</label>
					<input type="text" name="award_amount3" id="award_amount3" value='<%Response.write rs("award_amount3") %>'/>
					
					<label>Award Date</label>
					<input type="text" name="award_date3" id="award_date3" class="date" value='<%Response.write rs("award_date3") %>'/>
					<br/><br/><br/><br/>
					<br/><br/><br/><br/>
					</p>
				  
			
					

					
					<br/>
					
					<br/><br/>
					 <button type="submit" name="Submit" onclick="this.form.action='AfterEditStudent.asp?UIN=' + this.value; this.forms.submit();" value='<% Response.write rs("UIN") %>'>Save</button><br /><br />
					</fieldset>
				   
					
				</form>
				
			   </div>
			   
			</div>
			</div>
			<br/>
			<!--#include file="footer.asp"-->
</body>
</html>

