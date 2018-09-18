<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
    ErrMsg = Request("ErrMsg")
    UIN = Request("Button1")
set rs=Server.CreateObject("ADODB.recordset")
Query = "select * from CurrentPHDStudents where UIN ='" & UIN & "'"
    rs.Open query,conn
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | PHD Student Records</title>
<link rel="stylesheet" href="css/tabStyle.css" type="text/css" media="screen" />
<link rel="stylesheet" href="css/screen.css" />
<script type="text/javascript" src="jquery/jquery-1.9.0.js"></script>
<script type="text/javascript" src="jquery/jquery.validate.js"></script>
<script src="jquery/jquery.mask.js" type="text/javascript"></script>
<script type="text/javascript">
    $(document).ready(function () {
        $('.date').mask('00/00/0000');
        $('.homephone').mask('(000) 000-0000');
        $('.workphone').mask('(000) 000-0000 x00000');
        $('.iphone').mask('+000 000 000 000');
        $('.zip').mask('000000');
        $('.gpa').mask('0.00');
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
    <div id="content" align=center>
        <div id="wrapper">
          <div id="navigation" style="display: none;">
                    <ul>
                        <li><a href="#">Student Info</a></li>
                       
                        
                    </ul>
              </div>
            <div id="steps">
				<form id="studentForm" method="post">

					<h3>View PHD Student Information</h3>
                     <br/>
                    <a href=PHDCurrentStudents.asp>Back to Show Students</a> 
                    <br/> <br/>

                    <p>
                    <label>Student Info</label>
                    <br/><br/><br/><br/><br/>
					<label>First Name</label>
					<input type="text" name="FirstName" id="FirstName" value="<%Response.write rs("FirstName") %>" readonly=true/>   
                    <label>Middle Name</label>
					<input type="text" name="middlename" id="middlename" value="<%Response.Write rs("MiddleName") %>" readonly=true/> 
                    <label>Last Name</label>
					<input type="text" name="lastname" id="lastname" value="<%Response.Write rs("LastName") %>" readonly=true/>    
                    <br/><br/><br/><br/>
                    
                    <label>UIN</label>
					<input type="text" name="UIN" id="UIN" value='<%Response.write rs("UIN") %>' readonly=true/> 
                    <label>Date of Birth</label>
					<input type="text" name="dob" class="date" id="dob" value='<%Response.write rs("DateOfBirth") %>' readonly=true/>
                    <br/><br/><br/><br/>

                    <label>Race/Ethnicity</label>
					<input type="text" name="Race_ethinicity" id="Race_ethinicity" value='<%Response.write rs("Race_ethinicity") %>' readonly=true/>                                                           
                    <label>Race Description</label>
                    <input type="text" name="Race_desc" id="Race_desc" value='<%Response.write rs("Race_desc") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Salutation</label>
					<input type="text" name="Salutation" id="Salutation" value='<%Response.Write rs("Salutation") %>' readonly=true/>  
                    <label>Alternate Name</label>
					<input type="text" name="maidenname" id="maidenname" value='<%Response.Write rs("MaidenName") %>' readonly=true/>                  
                    <label>Gender</label>
					<input type="text" name="gender" id="gender" value='<%Response.Write rs("Gender") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    
                    <label>Current Address 1</label>
					<input type="text" name="mailingAddress1" id="mailingAddress1" value='<%Response.Write rs("MailingAddress1") %>' readonly=true/>
                    <label>Current Address 2</label>
					<input type="text" name="mailingAddress2" id="mailingAddress2" value='<%Response.Write rs("MailingAddress2") %>' readonly=true/>
                    <label>Current City</label>
					<input type="text" name="mailingcity" id="mailingcity" value='<%Response.Write rs("MailingCity") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Current ZipCode</label>
					<input type="text" name="mailingzipcode" class="zip" id="mailingzipcode" value='<%Response.Write rs("MailingZipCode") %>' readonly=true/>
                    <label>Current State</label>
					<input type="text" name="mailingstate" id="mailingstate" value='<%Response.Write rs("MailingState") %>' readonly=true/>
                    <label>Current Country</label>
					<input type="text" name="country" id="country" value='<%Response.Write rs("Country") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Home Phone</label>
					<input type="text" name="homephone" class="homephone" id="homephone" value='<%Response.Write rs("HomePhone") %>' readonly=true/>
                    <label>Cell Phone</label>
					<input type="text" name="cellphone" class="cellphone" id="cellphone" value='<%Response.Write rs("CellPhone") %>' readonly=true/>
                    <label>International Phone</label>
					<input type="text" name="internationalphonenumber" class="iphone" id="internationalphonenumber" value='<%Response.write rs("InternationalPhoneNumber") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>SO Name</label>
					<input type="text" name="SO_Name"  id="SO_Name" value='<%Response.write rs("SO_Name") %>'readonly=true/>
                    <label>Email</label>
					<input type="text" name="email" id="email" value='<%Response.write rs("Email") %>' readonly=true/>  
                    <label>FAX</label>
					<input type="text" name="fax"  id="fax" value='<%Response.write rs("Fax") %>' readonly=true/>                                                          
                    <br/><br/><br/><br />  
               
                    <label>UG College</label>
                    <input type="text" name="ugcollege" id="ugcollege" value='<%Response.write rs("UGCollege") %>' readonly=true/>                 
                    <label>UG GPA</label>
					<input type="text" name="uggpa" class="gpa" id="uggpa" value='<%Response.write rs("UGGPA") %>' readonly=true/>
                    <label>UG Major</label>
					<input type="text" name="ugmajor" id="ugmajor" value='<%Response.write rs("UGMajor") %>' readonly=true/>
                    <br/><br/><br/><br/>
                                        
                    <label>Grad College</label>
                    <input type="text" name="gradcollege" id="gradcollege" value='<%Response.write rs("GradCollege") %>' readonly=true/>
                    <label>Grad GPA</label>
					<input type="text" name="gradgpa" class="gpa" id="gradgpa" value='<%Response.write rs("GradGPA") %>' readonly=true/>                 
                    <label>Grad Major</label>
					<input type="text" name="gradmajor" id="gradmajor" value='<%Response.write rs("GradMajor") %>' readonly=true/>
                    <br/><br/><br/><br/>

                    <label>Program Type</label>
					<input type="text" name="ProgramType" id="ProgramType" value='<%Response.Write rs("Type") %>' readonly=true/>  
                    <label>Grad Degree</label>
					<input type="text" name="graddegree" id="graddegree" value='<%Response.write rs("GradDegree") %>' readonly=true/>
                    <br/><br/><br/><br/>
                        
                    <label>Date of Defense</label>
					<input type="text" name="DateofDefense" class ="date" id="DateofDefense" value='<%Response.Write rs("DateofDefense") %>' readonly=true/> 
                    <label>Preliminary Exam Date</label>
					<input type="text" name="DateofPreliminaryExam" class ="date" id="DateofPreliminaryExam" value='<%Response.Write rs("DateofPreliminaryExam") %>' readonly=true/> 
                    <label>Comprehensive Exam Date</label>
					<input type="text" name="DateofComprehensiveExam" class ="date" id="DateofComprehensiveExam" value='<%Response.Write rs("DateofComprehensiveExam") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Advisor</label>
					<input type="text" name="advisor" id="advisor" value='<%Response.Write rs("Advisor") %>' readonly=true/>
                    <label>Reason for Refusion</label>
					<input type="text" name="ReasonforRefusion" id="ReasonforRefusion" value='<%Response.write rs("ReasonforRefusion") %>' readonly=true/>
                    <label>Admit Term</label>
					<input type="text" name="AdmitTerm" id="AdmitTerm" value='<%Response.Write rs("AdmitTerm") %>' readonly=true/>
                    <br/><br/><br/><br/>

                    <label>Applying For Graduation</label>
					<input type="text" name="ApplyingForGraduation" id="ApplyingForGraduation" value='<%Response.Write rs("ApplyingforGraduation") %>' readonly=true/>
                    <label>Graduation Term Applied For</label>
					<input type="text" name="GraduationTermAppliedFor" id="GraduationTermAppliedFor" value='<%Response.Write rs("GraduationTermAppliedfor") %>' readonly=true/>
                    <label>Term Graduated</label>
					<input type="text" name="TermGraduated" id="TermGraduated" value='<%Response.write rs("TermGraduated") %>' readonly=true/>
                    <br/><br/><br/><br/>

                    <label>Entered By</label>
                    <input type="text" name="EnteredBy" id="EnteredBy" value='<%Response.write rs("EnteredBy") %>' readonly=true/>
                    <label>Date Entered</label>
                    <input type="text" name="DateEntered" class ="date" id="DateEntered" value='<%Response.write rs("DateEntered") %>' readonly=true/>
                    <label>Last Updated By</label>
                    <input type="text" name="LastUpdatedBy" id="LastUpdatedBy" value='<%Response.write rs("LastUpdatedBy") %>' readonly=true/>
                   
                    <br /><br /><br /><br />
          
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/>
                        </p>
					<button type="submit" name="Button1" onclick="this.form.action='EditPHDCurrentStudent.asp?UIN=' + this.value; this.forms.submit();" value=<% Response.write rs("UIN") %>>Edit Student</button><br /><br />
                    <button type="submit" name="Submit" onclick="this.form.action='RemovePHDStudent.asp?UIN=' + this.value; this.forms.submit();" value=<% Response.write rs("UIN") %>>Remove Student</button><br /><br />
					<br/><br/>
                  
                    
				</form>
               </div>
               <!--#include file="footer.asp"-->
			</div>
        </div>
            <br/>
</body>
</html>
