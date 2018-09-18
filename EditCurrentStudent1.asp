<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UIN = Request("UIN")
set rs=Server.CreateObject("ADODB.recordset")
query="select * from CurrentStudents where UIN ='"& UIN &"'"
rs.Open query,conn
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
      
                    <br />
        <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
        <br />
    <div id="wrapper">
     <div id="navigation" style="display: none;">
                    <ul>
                        <li><a href="#">Student Info</a></li>
                        <li><a href="#">MPH</a></li>
                        
                    </ul>
              </div>
    <div id="steps" align="center">
   
				<form id="studentForm" method="post" action="AfterEditCurrentStudent.asp?UIN=' + this.value; this.forms.submit();" value=<% Response.write rs("UIN") %>>
                    <fieldset class="step">
                <legend></legend>

                    <p>
                    
                    <br/><br/><br/>
					<label>First Name</label>
					<input type="text" name="firstname" required id="firstname" value="<%Response.write rs("FirstName") %>"/>   
                    <label>Middle Name</label>
					<input type="text" name="middlename" id="middlename" value="<%Response.write rs("middlename") %>"/> 
                    <label>Last Name</label>
					<input type="text" name="lastname" required id="lastname" value="<%Response.write rs("lastname") %>"/>    
                    <br/><br/><br/><br/>
                    
                    <label>UIN</label>
					<input type="text" name="UIN" required id="UIN" value='<%Response.write rs("UIN") %>' readonly=true/> 
                    <label>Date of Birth (MM/DD/YYYY)</label>
					<input type="text" name="dob" class="date" required id="dob" value='<%Response.write rs("DateOfBirth") %>' readonly=true/>
					<label>Degree Program</label>
					<input type="text" name="DegreeProgram" required id="DegreeProgram" value='<%Response.write rs("DegreeProgram") %>' /> 
                    <br/><br/><br/><br/>
                    
                    <label>Salutation</label>
					<input type="text" name="Salutation" value='<%Response.write rs("Salutation") %>' />  
                    <label>Alternate Name</label>
					<input type="text" name="maidenname" value="<%Response.write rs("maidenname") %>" />                  
                    <label>Gender</label>
                    <input type="text" name="gender" required id="gender" value='<%Response.write rs("gender") %>'  />
   	                                  
                    <br/><br/><br/><br/>
                    
                    <label>Current Address 1</label>
					<input type="text" name="currentAddress1" required id="currentAddress1" value="<%Response.write rs("currentAddress1") %>" />
                    <label>Current Address 2</label>
					<input type="text" name="currentAddress2" value="<%Response.write rs("currentAddress2") %>" />
                    <label>Current City</label>
					<input type="text" name="currentcity" required id="currentcity" value='<%Response.write rs("currentcity") %>' />
                    <br/><br/><br/><br/>
                    
                    <label>Current ZipCode</label>
					<input type="text" name="currentzipcode"  required id="currentzipcode" value='<%Response.write rs("currentzipcode") %>' />
                    <label>Current State</label>
                    <input type="text" name="currentstate" required id="currentstate" value='<%Response.write rs("currentstate") %>' />
                        
                    <label>Current Country</label>
					<input type="text" name="currentcountry" id="currentcountry" value='<%Response.write rs("currentcountry") %>' />
                    <br/><br/><br/><br/>
                    
                    <label>Home Phone</label>
					<input type="text" name="homephone" class="homephone" required id="homephone" value='<%Response.write rs("homephone") %>' />
                    <label>Work Phone</label>
					<input type="text" name="workphone" class="homephone" id="workphone" value='<%Response.write rs("workphone") %>' />
                    <label>International Phone</label>
					<input type="text" name="internationalphonenumber" class="homephone" id="internationalphonenumber" value='<%Response.write rs("InternationalPhoneNumber") %>' />
                    <br/><br/><br/><br/>
                    
                    <label>Cell Phone</label>
					<input type="text" name="cellphone" class="homephone" id="cellphone" value='<%Response.write rs("workphone") %>' />
                    <label>UIC Email</label>
					<input type="text" name="email" id="email" value='<%Response.write rs("email") %>' />
                    <label>Personal Email</label>
					<input type="text" name="Personalemail" id="Personalemail" value='<%Response.write rs("Personalemail") %>' />
                     
					<br/><br/><br/><br/>

                    <label>Race/Ethnicity</label>
					<input type="text" name="Race_ethinicity"  id="Race_ethinicity" value='<%Response.write rs("Race_ethinicity") %>' />                                                           
                    <label>Race SubCategory</label>
                    <input type="text" name="Race_SubCategory" id="Race_SubCategory" value='<%Response.write rs("Race_SubCategory") %>'/>
                    
					<label>Limited Status</label>
                    <select name="LimitedStatus" id="LimitedStatus">
                    <option value="<%= rs.Fields(29) %>"><%= rs.Fields(29) %></option>
                    <option value="Yes">Yes</option>
                    <option value="No">No</option>
                    </select>
                        <br/><br/><br/><br />

                    <label>Comments</label>
                    <textarea id="comments" name="comments" cols="45" rows="5" ><%Response.write rs("Comments") %></textarea>
                    <br/><br/><br/><br/>
                    
                        
                 
                        </p>
                        <button type="submit" name="Submit" onclick="this.form.action='AfterEditCurrentStudent1.asp?UIN=' + this.value; this.forms.submit();" value=<% Response.write rs("UIN") %>>Save Changes</button><br /><br />
                        </fieldset>
                   

                    <div id="application" align="center">
                    <fieldset class="step">
                    <legend></legend>
                
                    <p>
                    <br/>
                    
                    <label>Program Type</label>

					<select name="ProgramType" id="ProgramType">
                    <option value="<%= rs.Fields(30) %>"><%= rs.Fields(30) %></option>
                    <option value="FT">FT</option>
                    <option value="PM">PM</option>
                    <option value="Adv">Adv</option>
                    <option value="MPH-Adv">MPH-Adv</option>
                    <option value="MPH-FT">MPH-FT</option>
                    <option value="MPH-PM">MPH-PM</option>
                    <option value="TR-FT">TR-FT</option>
                    <option value="TR-PM">TR-PM</option>
                    </select>
                    <label>MPH Year</label>
                        
                    <select name="CurrentYear"  id="CurrentYear">
                    <option value="<%= rs.Fields(56) %>"><%= rs.Fields(56) %></option>
                    <option value=""></option>
                    <option value="Foundation">Foundation</option>
                    <option value="Concentration">Concentration</option>
                    <option value="Specialization">Specialization</option>
                    <option value="Generalist">Generalist</option>
                    <option value="Year1">Year1</option>
                    <option value="Year2">Year2</option>
                    <option value="Year3">Year3</option>
                    <option value="Year4">Year4</option>
                    
                    </select>  
					<label>Forward to Field</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="ForwardtoField"  id="ForwardtoField" class="checkboxField" value='<%Response.write rs("ForwardtoField") %>' />
                   
                         
                     <br/><br/><br/><br/>      
                    
					<label>Field Type</label>
                    <input type='hidden' style="margin:0;width:20px;height:20px;" class="cbhidden" value='<%Response.write rs("Field_Type") %>' />
                    <input type="checkbox" style="margin:0;width:20px;height:20px;" name="Field_Type" class="cbft" value='F'/>
                    <label class="clearWidth">Foundation</label>
                    <input type="checkbox" style="margin:0;width:20px;height:20px;" name="Field_Type"  class="cbft" value='C'/>
                    <label class="clearWidth">Concentration</label>
                    <label>Specialization</label>
                    <select name ="concentration" id="concentration">
                    <option value="<%= rs.Fields(31) %>"><%= rs.Fields(31) %></option>
                    <option value=""></option>
                    <option value="CHF">CHF</option>
                    <option value="CHUD">CHUD</option>
                    <option value="MH">MH</option>
                    <option value="SCH">SCH</option>
                    <option value="OCP">OCP</option>
                    </select>  
                     <label>Withdrawn</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="Withdrawn"  id="Withdrawn" class="checkboxField" value='<%Response.write rs("Withdrawn") %>' />
                    <br/><br/><br/><br/>
                    <label>Status</label>
                    <select name="Status" id="Status">
                    <option value="<%= rs.Fields(38) %>"><%= rs.Fields(38) %></option>
                    <option value=""> </option>
                    <option value="LOA">LOA</option>
                    <option value="DROP">DROP</option>
                    <option value="DENR">DENR</option>
                    <option value="TRANS">TRANS</option>
                    <option value="DNR">DNR</option>
                    <option value="PRO1">PRO1</option>
                    <option value="PRO2">PRO2</option>
                    <option value="GRAD">GRAD</option>
                    <option value="LTS">LTS</option>
                    <option value="WDN">WDN</option>
					</select>
                    <label>Withdrawn Date</label>
					<input type="text" name="WithdrawnDate" class="date" id="WithdrawnDate" value='<%Response.write rs("WithdrawnDate") %>' />      
                    <label>Withdrawn Term</label>
                    <select name="WithdrawnTerm" id="Select10">
                    <option value="<%= rs.Fields(65) %>"><%= rs.Fields(65) %></option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Fall 2018">Fall 2018</option>
                    <option value="Summer 2015">Summer 2015</option>
                    <option value="Summer 2016">Summer 2016</option>
                    <option value="Summer 2017">Summer 2017</option>
                    <option value="Summer 2018">Summer 2018</option>
                    </select>
                        <br/><br/><br/><br/>
                    

                    <label>Withdrawal Reason</label>
                    <textarea id="WithdrawalReason" name="WithdrawalReason" cols="45" rows="5" ><%Response.write rs("WithdrawalReason") %></textarea>
                    <br/><br/><br/><br />
                        
                    <label>Confirmed</label>
					<input type="text" name="Confirmed" required id="Confirmed" value='<%Response.write rs("Confirmed") %>' /> 
                    <label>Confirmed Date (MM/DD/YYYY)</label>
					<input type="text" name="ConfirmedDate" class="date"  id="ConfirmedDate" value='<%Response.write rs("ConfirmedDate") %>' /> 
                    <label>Admit Term</label>
                    <input type="text" name="Admit_Term" required id="Admit_Term" value='<%Response.write rs("AdmitTerm") %>' readonly=true/>
					
                    <br/><br/><br/><br/>
                    
                    <label>Advisor</label>
                    <select name="advisor"  id="advisor">
                    <option value="<%= rs.Fields(36) %>"><%= rs.Fields(36) %></option>
                    <option value="All">All</option>
                    <option value="Bonecutter">Bonecutter</option>
                    <option value="Butterfield">Butterfield</option>         
                    <option value="Coats">Coats</option>
                    <option value="D'Angelo">D'Angelo</option>
                    <option value="DeNard">DeNard</option>
                    <option value="Doyle">Doyle</option>
                    <option value="Geiger">Geiger</option>
                    <option value="Gottlieb">Gottlieb</option>
                    <option value="Hairston">Hairston</option>
                    <option value="Hounmenou">Hounmenou</option>
                    <option value="Hsieh">Hsieh</option>
                    <option value="Johnson">Johnson</option> 
                    <option value="Leathers">Leathers</option>
                    <option value="Lu">Lu</option>
                    <option value="McCoy">McCoy</option>
                    <option value="McLeod">McLeod</option>
                    <option value="Mitchell">Mitchell</option>
                    <option value="Salvadore">Salvadore</option>
                    <option value="Swartz">Swartz</option>
                    <option value="Watson">Watson</option>
                    <option value="Wilson">Wilson</option>                   
                    </select>  
					
                    <label>Track</label>
                    <input type="text" name="Track" id="Track" value='<%Response.write rs("Track") %>' />               
		            <label>Decision</label>
					<input type="text" name="Decision" id="Decision" value='<%Response.write rs("Decision") %>' />  
                    
                    <br/><br/><br/><br/>

                    <label>Applying For Graduation</label>
                    <select name="ApplyingForGraduation" id="ApplyingForGraduation">
                    <option value="<%= rs.Fields(39) %>"><%= rs.Fields(39) %></option>
                    <option value="Yes">Yes</option>
                    <option value="No">No</option>
                    </select>
                    <label>Graduation Term Applied For</label>
                    <select name="GraduationTermAppliedFor" id="Select1">
                    <option value="<%= rs.Fields(40) %>"><%= rs.Fields(40) %></option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Fall 2018">Fall 2018</option>
                    <option value="Summer 2015">Summer 2015</option>
                    <option value="Summer 2016">Summer 2016</option>
                    <option value="Summer 2017">Summer 2017</option>
                    <option value="Summer 2018">Summer 2018</option>
                    </select>  
					
                    <label>Term Graduated</label>
                    <select name="TermGraduated" id="TermGraduated">
                    <option value="<%= rs.Fields(41) %>"><%= rs.Fields(41) %></option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Fall 2018">Fall 2018</option>
                    <option value="Summer 2015">Summer 2015</option>
                    <option value="Summer 2016">Summer 2016</option>
                    <option value="Summer 2017">Summer 2017</option>
                    <option value="Summer 2018">Summer 2018</option>
                    </select>
             
					
                    <br/><br/><br/><br/><br/>

                    <label>Degree Applying For</label>
					<input type="text" name="DegreeApplyingFor"  id="DegreeApplyingFor" value='<%Response.write rs("DegreeApplyingFor") %>' />
                   
                    <label style="padding-left:40px">Graduated</label>
                    <input type="checkbox" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="Graduated" id="Graduated" class="checkboxField" value='<%Response.write rs("Graduated") %>' />
                    <label>Graduation Date (MM/DD/YYYY)</label>
					<input type="text" name="GraduatedDate" class="date"  id="GraduatedDate" value='<%Response.write rs("GraduatedDate") %>' /> 
                    <br/><br/><br/><br/>

                    
                    <label>Probation Start Term</label>
                    <select name="ProbationStartTerm" id="ProbationStartTerm">
                    <option value="<%= rs.Fields(48) %>"><%= rs.Fields(48) %></option>
                    <option value=""> </option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Fall 2018">Fall 2018</option>
                    </select> 
                    <label>Probation End Term</label>
                    <select name="ProbationEndTerm" id="ProbationEndTerm">
                    <option value="<%= rs.Fields(49) %>"><%= rs.Fields(49) %></option>
                    <option value=""> </option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Fall 2018">Fall 2018</option>
                    </select> 
                    <br/><br/><br/><br/>

                    <label>Leave of Absence Start Term</label>
                    <select name="LeaveofAbsenceStartTerm" id="LeaveofAbsenceStartTerm">
                    <option value="<%= rs.Fields(50) %>"><%= rs.Fields(50) %></option>
                    <option value=""> </option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Fall 2018">Fall 2018</option>
                    </select> 
                    <label>Leave of Absence End Term</label>
                    <select name="LeaveofAbsenceEndTerm" id="LeaveofAbsenceEndTerm">
                    <option value="<%= rs.Fields(51) %>"><%= rs.Fields(51) %></option>
                    <option value=""> </option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Fall 2018">Fall 2018</option>
                    </select>
                    <label>Modified Plan</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;margin-right: 140px;" name="ModifiedPlan"  id="ModifiedPlan" class="checkboxField" value='<%Response.write rs("ModifiedPlan") %>' />
                     
                    <br/><br/><br/><br/>
                    <label style="padding-left:40px"> IBHE/Evidence-Based MH Practice w/Children</label>
                    <input type="checkbox"  style="margin:0;width:20px;height:20px;margin-right: 140px;" name="IBHE_Certificate" id="IBHE_Certificate" class="checkboxField" value="<%Response.write rs("IBHE_Certificate") %>" />
                    <label>Certificate Start Term</label>
					<select name="Certificate_StartTerm" id="Certificate_StartTerm">
                    <option value="<%= rs.Fields(58) %>"><%= rs.Fields(58) %></option>
                    <option value=""> </option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Fall 2018">Fall 2018</option>
                    </select>
                    <br/><br/><br/><br/><br/><br/>
                    <label style="padding-left:40px"> Child Welfare Traineeship Project</label>
                    <input type="checkbox"  style="margin:0;width:20px;height:20px;margin-right: 140px;" name="ChildWelfareTraineeshipProject" id="ChildWelfareTraineeshipProject" class="checkboxField" value="<%Response.write rs("ChildWelfareTraineeshipProject") %>" />
                    <label>Child Welfare Traineeship Project Start Term</label>
					<select name="ChildWelfareTraineeshipProjectStartTerm" id="ChildWelfareTraineeshipProjectStartTerm">
                    <option value="<%= rs.Fields(60) %>"><%= rs.Fields(60) %></option>
                    <option value=""> </option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Fall 2018">Fall 2018</option>
                    </select>
                    <label>Confirmed Due Date</label>
					<input type="text" name="ConfirmedDueDate" class="date"  id="ConfirmedDueDate" value='<%Response.write rs("ConfirmedDueDate") %>' /> 
                    <br/><br/><br/><br/>
                  
                        </p>
                        <button type="submit" name="Submit" onclick="this.form.action='AfterEditCurrentStudent1.asp?UIN=' + this.value; this.forms.submit();" value=<% Response.write rs("UIN") %>>Save Changes</button><br /><br />
                        </fieldset>
                       

					<br/><br/>
                      
                    </form>
			
                            </div>
        
        </div>
      </div>
        
               <!--#include file="footer.asp"-->
			
       
     

            <br/>
</body>
</html>