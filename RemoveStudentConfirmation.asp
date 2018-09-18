<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UIN = Request("Submit")
    set rs=Server.CreateObject("ADODB.recordset")
     query="select * from CurrentStudents where UIN ='"& UIN &"'"
rs.Open query,conn
   ProgramType = rs("ProgramType")

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | Student Records</title>
<link rel="stylesheet" href="css/tabStyle.css" type="text/css" media="screen" />
<link rel="stylesheet" href="css/screen.css" />
<script type="text/javascript" src="jquery/jquery-1.9.0.js"></script>
<script type="text/javascript" src="jquery/jquery.validate.js"></script>
<script type="text/javascript" src="jquery/sliding.form.js"></script>
<script src="jquery/jquery.mask.js" type="text/javascript"></script>
<script type="text/javascript" src="jquery/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
    $(document).ready(function () {
        $('.date').mask('00/00/0000');
        $('.homephone').mask('(000) 000-0000');
        $('.workphone').mask('(000) 000-0000 x00000');
        $('.iphone').mask('+000 000 000 000');
       
        $('.gpa').mask('0.00');

        var fieldval = $('input.cbhidden').val();
        if (fieldval != '' || fieldval != undefined) {
            $('input.cbft[value=' + fieldval + ']').attr('checked', 'checked');
        }

        $('input.rbfield').removeAttr('checked');
        var checkedElm = $('input.rbfieldhidden').val();
        if (checkedElm != '' || checkedElm != undefined) {
            $('input:radio[value=' + checkedElm + ']').attr('checked', 'checked');
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
    <h3>View Student Information</h3>
                     <br/>
                    <a href="ShowCurrentStudents.asp">Back to Show Students</a>  
                    <br/> <br/>
        <% If ProgramType = "MPH" or ProgramType = "MPH-Adv" or ProgramType = "MPH-FT" or ProgramType = "MPH-PM" Then %>
                   
                    
    <div id="wrapper">
     <div id="navigation" style="display: none;">
                    <ul>
                        <li><a href="#">Student Info</a></li>
                        <li><a href="#">MPH</a></li>
                        
                    </ul>
              </div>
    <div id="steps" align="center">
   
				<form id="studentForm" method="post" action="RemoveStudent.asp?UIN=' + this.value; " value='<% Response.write rs("UIN") %>'>
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
                    <br/><br/><br/><br/>
                    
                    <label>UIN</label>
					<input type="text" name="UIN" required id="UIN" value='<%Response.write rs("UIN") %>' readonly=true/> 
                    <label>Date of Birth</label>
					<input type="text" name="dob" class="date" required id="dob" value='<%Response.write rs("DateOfBirth") %>' readonly=true/>
					<label>Degree Program</label>
					<input type="text" name="DegreeProgram" required id="DegreeProgram" value='<%Response.write rs("DegreeProgram") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    
                    <label>Salutation</label>
					<input type="text" name="Salutation" required id="Salutation" value='<%Response.write rs("Salutation") %>' readonly=true/>  
                    <label>Alternate Name</label>
					<input type="text" name="maidenname" required id="maidenname" value='<%Response.write rs("maidenname") %>' readonly=true/>                  
                    <label>Gender</label>
					<input type="text" name="gender" required id="gender" value='<%Response.write rs("gender") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    
                    <label>Current Address 1</label>
					<input type="text" name="currentAddress1" required id="currentAddress1" value='<%Response.write rs("currentAddress1") %>' readonly=true/>
                    <label>Current Address 2</label>
					<input type="text" name="currentAddress2" required id="currentAddress2" value='<%Response.write rs("currentAddress2") %>' readonly=true/>
                    <label>Current City</label>
					<input type="text" name="currentcity" required id="currentcity" value='<%Response.write rs("currentcity") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Current ZipCode</label>
					<input type="text" name="currentzipcode" class="zip" required id="currentzipcode" value='<%Response.write rs("currentzipcode") %>' readonly=true/>
                    <label>Current State</label>
					<input type="text" name="currentstate" required id="currentstate" value='<%Response.write rs("currentstate") %>' readonly=true/>
                    <label>Current Country</label>
					<input type="text" name="currentcountry" required id="currentcountry" value='<%Response.write rs("currentcountry") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Home Phone</label>
					<input type="text" name="homephone" class="homephone" required id="homephone" value='<%Response.write rs("homephone") %>' readonly=true/>
                    <label>Work Phone</label>
					<input type="text" name="workphone" class="workphone" required id="workphone" value='<%Response.write rs("workphone") %>' readonly=true/>
                    <label>International Phone</label>
					<input type="text" name="internationalphonenumber" class="iphone" id="internationalphonenumber" value='<%Response.write rs("InternationalPhoneNumber") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Cell Phone</label>
					<input type="text" name="cellphone" class="cellphone"  id="cellphone" value='<%Response.write rs("workphone") %>' readonly=true/>
                        <label>UIC Email</label>
					<input type="text" name="email" id="email" value='<%Response.write rs("email") %>' readonly=true/>
                    <label>Personal Email</label>
					<input type="text" name="Personalemail" id="Personalemail" value='<%Response.write rs("Personalemail") %>' readonly=true/>
                     
					<br/><br/><br/><br/>

                    <label>Race/Ethnicity</label>
					<input type="text" name="Race_ethinicity"  id="Race_ethinicity" value='<%Response.write rs("Race_ethinicity") %>' readonly=true/>                                                           
                    <label>Race SubCategory</label>
                    <input type="text" name="Race_SubCategory" id="Race_SubCategory" value='<%Response.write rs("Race_SubCategory") %>' readonly=true/>
                    <label>Limited Status</label>
					<input type="text" name="LimitedStatus" required id="LimitedStatus" value='<%Response.write rs("LimitedStatus") %>' readonly=true/>
                       <br/><br/><br/><br />

                    <label>Comments</label>
                    <textarea id="comments" name="comments" cols="45" rows="5" readonly=true><%Response.write rs("Comments") %></textarea>
                    <br/><br/><br/><br/>
                   
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/>
                        </p>
            <button type="submit" name="Submit" onclick="this.form.action='RemoveStudent.asp?UIN=' + this.value; this.forms.submit();" value='<% Response.write rs("UIN") %>'>Confirm to Remove Student</button><br /><br />

                        </fieldset>

                    <div id="application" align="center">
                    <fieldset class="step">
                    <legend></legend>
                
                    <p>
                    <br/>
                    
                    <label>Program Type</label>
					<input type="text" name="ProgramType" required id="ProgramType" value='<%Response.write rs("ProgramType") %>' readonly=true/>  
                    <label>MPH Year</label>
					<input type="text" name="CurrentYear" required id="CurrentYear" value='<%Response.write rs("CurrentYear") %>' readonly=true/>
                    <label style="padding-left:40px">Forward to Field</label>
                    <input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="ForwardtoField" id="ForwardtoField" class="checkboxField" value='<%Response.write rs("ForwardtoField") %>' />
                    <br/><br/><br/><br/>

                    <label>Field Type</label>
                    <input type='hidden' style="margin:0;width:20px;height:20px;" class="cbhidden" value='<%Response.write rs("Field_Type") %>' />
                    <input type="checkbox" style="margin:0;width:20px;height:20px;" name="Field_Type" disabled="disabled" class="cbft" value='F'/>
                    <label class="clearWidth">Foundation</label>
                    <input type="checkbox" style="margin:0;width:20px;height:20px;" name="Field_Type" disabled="disabled" class="cbft" value='C'/>
                    <label class="clearWidth">Concentration</label>
				    <label>Concentration</label>
                    <input type="text" name="concentration" id="concentration" value='<%Response.write rs("Concentration") %>' readonly="true"/>  
                       
                    
                    <br/><br/><br/><br/>
                        
                    <label>Confirmed</label>
					<input type="text" name="Confirmed" required id="Confirmed" value='<%Response.write rs("Confirmed") %>' readonly=true/> 
                    <label>Confirmed Date</label>
					<input type="text" name="ConfirmedDate" class="date" required id="ConfirmedDate" value='<%Response.write rs("ConfirmedDate") %>' readonly=true/> 
                    <label>Admit Term</label>
					<input type="text" name="AdmitTerm" required id="AdmitTerm" value='<%Response.write rs("AdmitTerm") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Advisor</label>
					<input type="text" name="advisor" required id="advisor" value='<%Response.write rs("advisor") %>' readonly=true/>
                    <label>Track</label>
					<input type="text" name="Track" required id="Track" value='<%Response.write rs("Track") %>' readonly=true/>
                   <label>Decision</label>
					<input type="text" name="Decision" required id="Decision" value='<%Response.write rs("Decision") %>' readonly=true/>   
                    <br/><br/><br/><br/>

                    <label>Applying For Graduation</label>
					<input type="text" name="ApplyingForGraduation" required id="ApplyingForGraduation" value='<%Response.write rs("ApplyingForGraduation") %>' readonly=true/>
                    <label>Graduation Term Applied For</label>
					<input type="text" name="GraduationTermAppliedFor" required id="GraduationTermAppliedFor" value='<%Response.write rs("GraduationTermAppliedFor") %>' readonly=true/>
                    <label>Term Graduated</label>
					<input type="text" name="TermGraduated" required id="TermGraduated" value='<%Response.write rs("TermGraduated") %>' readonly=true/>
                    <br/><br/><br/><br/><br/>

                    <label>Degree Applying For</label>
					<input type="text" name="DegreeApplyingFor" required id="DegreeApplyingFor" value='<%Response.write rs("DegreeApplyingFor") %>' readonly=true/>
                    
					<label style="padding-left:40px">Graduated</label>
                    <input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="Graduated" id="Graduated" class="checkboxField" value='<%Response.write rs("Graduated") %>' />
                    
                    <label>Graduation Date</label>
					<input type="text" name="GraduatedDate" class="date" required id="GraduatedDate" value='<%Response.write rs("GraduatedDate") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    
                    <label>Probation Start Term</label>
					<input type="text" name="ProbationStartTerm"  id="ProbationStartTerm" value='<%Response.write rs("ProbationStartTerm") %>' readonly=true/> 
                    <label>Probation End Term</label>
					<input type="text" name="ProbationEndTerm"  id="ProbationEndTerm" value='<%Response.write rs("ProbationEndTerm") %>' readonly=true/> 
                    <br/><br/><br/><br/>

                    <label>Leave of Absence Start Term</label>
					<input type="text" name="LeaveofAbsenceStartTerm"  id="LeaveofAbsenceStartTerm" value='<%Response.write rs("LeaveofAbsenceStartTerm") %>' readonly=true/> 
                    <label>Leave of Absence End Term</label>
					<input type="text" name="LeaveofAbsenceEndTerm"  id="LeaveofAbsenceEndTerm" value='<%Response.write rs("LeaveofAbsenceEndTerm") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    <label style="padding-left:40px">IBHE Certificate</label>
                    <input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="IBHE_Certificate" id="IBHE_Certificate" class="checkboxField" value='<%Response.write rs("IBHE_Certificate") %>' />
                    <label>Certificate Start Term</label>
					<input type="text" name="Certificate_StartTerm"  id="Certificate_StartTerm" value='<%Response.write rs("Certificate_StartTerm") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/>
                        </p>

                        
                    <button type="submit" name="Submit" onclick="this.form.action='RemoveStudent.asp?UIN=' + this.value; this.forms.submit();" value='<% Response.write rs("UIN") %>'>Confirm to Remove Student</button><br /><br />

                        </fieldset>
                        </div>  
                        </form>
        </div>
        </div>
                        
        <%Else %>
                     

                         <div id="wrapper">
     <div id="navigation" style="display: none;">
                    <ul>
                        <li><a href="#">Student Info</a></li>
                    </ul>
              </div>
    <div id="steps" align="center">
   
				<form id="studentform" method="post" action="RemoveStudent.asp?UIN=' + this.value; " value=<% Response.write rs("UIN") %>>
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
                    <br/><br/><br/><br/>
                    
                    <label>UIN</label>
					<input type="text" name="UIN" required id="UIN" value='<%Response.write rs("UIN") %>' readonly=true/> 
                    <label>Date of Birth</label>
					<input type="text" name="dob" class="date" required id="dob" value='<%Response.write rs("DateOfBirth") %>' readonly=true/>
					<label>Degree Program</label>
					<input type="text" name="DegreeProgram" required id="DegreeProgram" value='<%Response.write rs("DegreeProgram") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    
                    <label>Salutation</label>
					<input type="text" name="Salutation" required id="Salutation" value='<%Response.write rs("Salutation") %>' readonly=true/>  
                    <label>Alternate Name</label>
					<input type="text" name="maidenname" required id="maidenname" value='<%Response.write rs("maidenname") %>' readonly=true/>                  
                    <label>Gender</label>
					<input type="text" name="gender" required id="gender" value='<%Response.write rs("gender") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    
                    <label>Current Address 1</label>
					<input type="text" name="currentAddress1" required id="currentAddress1" value='<%Response.write rs("currentAddress1") %>' readonly=true/>
                    <label>Current Address 2</label>
					<input type="text" name="currentAddress2" required id="currentAddress2" value='<%Response.write rs("currentAddress2") %>' readonly=true/>
                    <label>Current City</label>
					<input type="text" name="currentcity" required id="currentcity" value='<%Response.write rs("currentcity") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Current ZipCode</label>
					<input type="text" name="currentzipcode" class="zip" required id="currentzipcode" value='<%Response.write rs("currentzipcode") %>' readonly=true/>
                    <label>Current State</label>
					<input type="text" name="currentstate" required id="currentstate" value='<%Response.write rs("currentstate") %>' readonly=true/>
                    <label>Current Country</label>
					<input type="text" name="currentcountry" required id="currentcountry" value='<%Response.write rs("currentcountry") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Home Phone</label>
					<input type="text" name="homephone" class="homephone" required id="homephone" value='<%Response.write rs("homephone") %>' readonly=true/>
                    <label>Work Phone</label>
					<input type="text" name="workphone" class="workphone" required id="workphone" value='<%Response.write rs("workphone") %>' readonly=true/>
                    <label>International Phone</label>
					<input type="text" name="internationalphonenumber" class="iphone" id="internationalphonenumber" value='<%Response.write rs("InternationalPhoneNumber") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Cell Phone</label>
					<input type="text" name="cellphone" required id="cellphone" value='<%Response.write rs("workphone") %>' readonly=true/>
                        <label>UIC Email</label>
					<input type="text" name="email" id="email" value='<%Response.write rs("email") %>' readonly=true/>
                    <label>Limited Status</label>
					<input type="text" name="LimitedStatus" required id="LimitedStatus" value='<%Response.write rs("LimitedStatus") %>' readonly=true/>
                     
					<br/><br/><br/><br/>

                    <label>Race/Ethnicity</label>
					<input type="text" name="Race_ethinicity"  id="Race_ethinicity" value='<%Response.write rs("Race_ethinicity") %>' readonly=true/>                                                           
                    <label>Race SubCategory</label>
                    <input type="text" name="Race_SubCategory" id="Race_SubCategory" value='<%Response.write rs("Race_SubCategory") %>' readonly=true/>
                    <label>Personal Email</label>
					<input type="text" name="Personalemail" id="Personalemail" value='<%Response.write rs("Personalemail") %>' readonly=true/>
                        <br/><br/><br/><br />


                    <label>Comments</label>
                    <textarea id="comments" name="comments" cols="45" rows="5" readonly=true><%Response.write rs("Comments") %></textarea>
                    <br/><br/><br/><br/>
                    
                    <label>Program Type</label>
					<input type="text" name="ProgramType" required id="ProgramType" value='<%Response.write rs("ProgramType") %>' readonly=true/>  
                    <label>Current Year</label>
					<input type="text" name="CurrentYear" required id="CurrentYear" value='<%Response.write rs("CurrentYear") %>' readonly=true/>
                    <label style="padding-left:40px">Forward to Field</label>
                    <input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="ForwardtoField" id="ForwardtoField" class="checkboxField" value='<%Response.write rs("ForwardtoField") %>' />
                    <br/><br/><br/><br/>

                    <label>Field Type</label>
                    <input type='hidden' style="margin:0;width:20px;height:20px;" class="cbhidden" value='<%Response.write rs("Field_Type") %>' />
                    <input type="checkbox" style="margin:0;width:20px;height:20px;" name="Field_Type" disabled="disabled" class="cbft" value='F'/>
                    <label class="clearWidth">Foundation</label>
                    <input type="checkbox" style="margin:0;width:20px;height:20px;" name="Field_Type" disabled="disabled" class="cbft" value='C'/>
                    <label class="clearWidth">Concentration</label>
					<label>Concentration</label>
                    <input type="text" name="concentration" id="concentration" value='<%Response.write rs("Concentration") %>' readonly="true"/>   
                       
                    
                    <br/><br/><br/><br/>
                        
                    <label>Confirmed</label>
					<input type="text" name="Confirmed" required id="Confirmed" value='<%Response.write rs("Confirmed") %>' readonly=true/> 
                    <label>Confirmed Date</label>
					<input type="text" name="ConfirmedDate" class="date" required id="ConfirmedDate" value='<%Response.write rs("ConfirmedDate") %>' readonly=true/> 
                    <label>Admit Term</label>
					<input type="text" name="AdmitTerm" required id="AdmitTerm" value='<%Response.write rs("AdmitTerm") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Advisor</label>
					<input type="text" name="advisor" required id="advisor" value='<%Response.write rs("advisor") %>' readonly=true/>
                    <label>Track</label>
					<input type="text" name="Track" required id="Track" value='<%Response.write rs("Track") %>' readonly=true/>
                   <label>Decision</label>
					<input type="text" name="Decision" required id="Decision" value='<%Response.write rs("Decision") %>' readonly=true/>   
                    <br/><br/><br/><br/>

                    <label>Applying For Graduation</label>
					<input type="text" name="ApplyingForGraduation" required id="ApplyingForGraduation" value='<%Response.write rs("ApplyingForGraduation") %>' readonly=true/>
                    <label>Graduation Term Applied For</label>
					<input type="text" name="GraduationTermAppliedFor" required id="GraduationTermAppliedFor" value='<%Response.write rs("GraduationTermAppliedFor") %>' readonly=true/>
                    <label>Term Graduated</label>
					<input type="text" name="TermGraduated" required id="TermGraduated" value='<%Response.write rs("TermGraduated") %>' readonly=true/>
                    <br/><br/><br/><br/><br/>

                    <label>Degree Applying For</label>
					<input type="text" name="DegreeApplyingFor" required id="DegreeApplyingFor" value='<%Response.write rs("DegreeApplyingFor") %>' readonly=true/>
                    
					<label style="padding-left:40px">Graduated</label>
                    <input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="Graduated" id="Graduated" class="checkboxField" value='<%Response.write rs("Graduated") %>' />
                    
                    <label>Graduation Date</label>
					<input type="text" name="GraduatedDate" class="date" required id="GraduatedDate" value='<%Response.write rs("GraduatedDate") %>' readonly=true/> 
                    <br/><br/><br/><br/>

                    <label>Probation Start Term</label>
					<input type="text" name="ProbationStartTerm"  id="ProbationStartTerm" value='<%Response.write rs("ProbationStartTerm") %>' readonly=true/> 
                    <label>Probation End Term</label>
					<input type="text" name="ProbationEndTerm"  id="ProbationEndTerm" value='<%Response.write rs("ProbationEndTerm") %>' readonly=true/> 
                    <br/><br/><br/><br/>

                    <label>Leave of Absence Start Term</label>
					<input type="text" name="LeaveofAbsenceStartTerm"  id="LeaveofAbsenceStartTerm" value='<%Response.write rs("LeaveofAbsenceStartTerm") %>' readonly=true/> 
                    <label>Leave of Absence End Term</label>
					<input type="text" name="LeaveofAbsenceEndTerm"  id="LeaveofAbsenceEndTerm" value='<%Response.write rs("LeaveofAbsenceEndTerm") %>' readonly=true/> 
                    <br/><br/><br/><br/>

                    <label style="padding-left:40px">IBHE/Evidence-Based MH Practice w/Children</label>
                    <input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="IBHE_Certificate" id="IBHE_Certificate" class="checkboxField" value='<%Response.write rs("IBHE_Certificate") %>' />
                    <label>Certificate Start Term</label>
					<input type="text" name="Certificate_StartTerm"  id="Certificate_StartTerm" value='<%Response.write rs("Certificate_StartTerm") %>' readonly=true/> 
                    <br/><br/><br/><br/><br/><br/>
                    <label style="padding-left:40px">Child Welfare Traineeship Project</label>
                    <input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="ChildWelfareTraineeshipProject" id="ChildWelfareTraineeshipProject" class="checkboxField" value='<%Response.write rs("ChildWelfareTraineeshipProject") %>' />
                    <label>Child Welfare Traineeship Project Start Term</label>
					<input type="text" name="ChildWelfareTraineeshipProjectStartTerm"  id="ChildWelfareTraineeshipProjectStartTerm" value='<%Response.write rs("ChildWelfareTraineeshipProjectStartTerm") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/>
                        </p>
                       
                    <button type="submit" name="Submit" onclick="this.form.action='RemoveStudent.asp?UIN=' + this.value; this.forms.submit();" value='<% Response.write rs("UIN") %>'>Confirm to Remove Student</button><br /><br />

                        </fieldset>
                
					<br/><br/>
                  
                   </form> 
				</div>
                
					 
        </div>
         <%End If %>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</body>
</html>
