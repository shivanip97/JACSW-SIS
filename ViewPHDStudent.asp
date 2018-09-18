<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UIN = Request("Button1")
set rs=Server.CreateObject("ADODB.recordset")
query="select * from PHDApplicants where UIN ='"& UIN &"'"
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
    <h3>View PHD Student Information</h3>
                     <br/>
                    <a href="PHDApplication.asp?ID=220168">Back to PHD Students</a> 
                    <br/> <br/>
                    <% If Session("Username") = "test_ap" or Session("Username") = "cstoakl" or Session("Username") = "tmorri3" Then %>
                    <button style="border:none;outline:none;-moz-border-radius: 10px;-webkit-border-radius: 10px;font-weight:bold;margin: 0px auto;clear:both;padding: 7px 25px;font-size:22px;display: block; background:#4797ED;font-family:Century Gothic, Helvetica, sans-serif;" type="submit" name="Button1" onclick="studentForm.action='EditPHDStudent.asp?UIN=' + this.value; studentForm.submit();" value='<% Response.write rs("UIN") %>'>Edit Student</button><br /><br />
                    <% End If %>
    <div id="wrapper">
     <div id="navigation" style="display: none;">
                    <ul>
                        <li><a href="#">Student Demographics</a></li>
                        <li><a href="#">Application Info</a></li>
                        <li><a href="#">Comments</a></li>
                    </ul>
              </div>
    <div id="steps" align="center">
				<form id="studentForm" method="post" action="EditStudent.asp">
                
                    
                
              
                <fieldset class="step">
                <legend></legend>

                <p>
                    <br/><br/><br/>
                    <label>Last Name</label>
					<input type="text" name="lastname"  id="lastname" value='<%Response.write rs("lastname") %>' readonly=true/>
        
                    <label>UIN</label>
					<input type="text" name="uin"  id="uin" value='<%Response.write rs("uin") %>' readonly=true/> 

                    <br/><br/><br/><br />
                    <label>First Name</label>
					<input type="text" name="FirstName"  id="FirstName" value='<%Response.write rs("FirstName") %>' readonly=true/>   
                    <label>Middle Name</label>
					<input type="text" name="middlename" id="middlename" value='<%Response.write rs("middlename") %>' readonly=true/> 
                    <label>Salutation</label>
					<input type="text" name="Salutation" id="Salutation" value='<%Response.write rs("Salutation") %>'readonly=true />
                      
                    <br/><br/><br/><br />
                    <label>Maiden Name</label>
                    <input type="text" name="maidenname" id="maidenname" value='<%Response.write rs("maidenname") %>'/> 
                    <label>Gender</label>
					<input type="text" name="gender"  id="gender" value='<%Response.write rs("gender") %>' readonly=true/>
                    
                    <label>Date of Birth</label>
					<input type="text" name="dob" class="date" id="dob" value='<%Response.write rs("DateOfBirth") %>' readonly=true/>
                    <br/><br/><br/><br />
                    <label>Race/Ethnicity</label>
					<input type="text" name="Race_ethinicity" id="Race_ethinicity" value='<%Response.write rs("Race_ethinicity") %>' readonly=true/>                                                           
                    <label>Race Description</label>
                    <input type="text" name="Race_desc" id="Race_desc" value='<%Response.write rs("Race_desc") %>' readonly=true/>
                    <br/><br/><br/><br />
                    
                    
                    <label>SO Name</label>
					<input type="text" name="SO_Name"  id="SO_Name" value='<%Response.write rs("SO_Name") %>'readonly=true/>
                    <label>Email</label>
					<input type="text" name="email" id="email" value='<%Response.write rs("email") %>' readonly=true/>  
                    <label>FAX</label>
					<input type="text" name="fax"  id="fax" value='<%Response.write rs("fax") %>' readonly=true/>                                                          
                    <br/><br/><br/><br />               
                    <label>International Phone</label>
					<input type="text" name="InternationalPhoneNumber"  id="InternationalPhoneNumber" value='<%Response.write rs("InternationalPhoneNumber") %>' readonly=true/>
                    <label>Home Phone</label>
					<input type="text" name="homephone" class="homephone"  id="homephone" value='<%Response.write rs("homephone") %>' readonly=true/>
                    <label>Cell Phone</label>
					<input type="text" name="workphone" class="workphone" id="workphone" value='<%Response.write rs("workphone") %>' readonly=true/>                             
                    
                    <br/><br/><br/><br />
                    <label>Mailing Address 1</label>
					<input type="text" name="currentAddress1"  id="currentAddress1" value='<%Response.write rs("currentAddress1") %>' readonly=true/>
                    <label>Mailing Address 2</label>
					<input type="text" name="currentAddress2" id="currentAddress2" value='<%Response.write rs("currentAddress2") %>' readonly=true/>
                    <label>Mailing City</label>
					<input type="text" name="currentcity"  id="currentcity" value='<%Response.write rs("currentcity") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Mailing State</label>
                    <input type="text" name="currentstate"  id="currentstate" value='<%Response.write rs("currentstate") %>' readonly=true/>
                    
                    <label>Mailing Zip</label>
					<input type="text" name="currentzip" class="zip"  id="currentzip" value='<%Response.write rs("currentzipcode") %>'readonly=true/>
                    <label>Country</label>
					<input type="text" name="currentcountry"  id="currentcountry" value='<%Response.write rs("currentcountry") %>'readonly=true/>
                    

                    <br /><br /><br /><br />
                    
                                       
                  
                    <label>Citizenship Status</label>
                    <input type="text" name="Citizenship_Status" id="Citizenship_Status" value='<%Response.write rs("Citizenship_Status") %>' readonly=true/>
                    <label style="width:170px">Country of Citizenship</label>
					<input type="text" name="Country_of_Citizenship" id="Country_of_Citizenship" value='<%Response.write rs("Country_of_Citizenship") %>' readonly=true/>

                    <br/><br/><br/><br/>
                
                    <label>UG College</label>
                    <input type="text" name="ugcollege" id="ugcollege" value='<%Response.write rs("ugcollege") %>' readonly=true/>
                    
                    <label>UG GPA</label>
					<input type="text" name="uggpa" class="gpa" id="uggpa" value='<%Response.write rs("UGGPA") %>' readonly=true/>
                    <label>UG Major</label>
					<input type="text" name="ugmajor" id="ugmajor" value='<%Response.write rs("UGMajor") %>' readonly=true/>
                    <br/><br/><br/><br/>
                                        
                    <label>Grad College</label>
                    <input type="text" name="gradcollege" id="gradcollege" value='<%Response.write rs("gradcollege") %>' readonly=true/>
		
                    <label>Grad GPA</label>
					<input type="text" name="gradgpa" class="gpa" id="gradgpa" value='<%Response.write rs("GradGPA") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    
                    <label>Grad Major</label>
					<input type="text" name="gradmajor" id="gradmajor" value='<%Response.write rs("GradMajor") %>' readonly=true/>
                    <label>Grad Degree</label>
					<input type="text" name="graddegree" id="graddegree" value='<%Response.write rs("GradDegree") %>' readonly=true/>
                    
                    <br/><br/><br/><br/>
                    
                    </p>
                    </fieldset>

                    <div id="application" align="center">
                    <fieldset class="step">
                    <legend></legend>
                
                    <p>
                    <br/>
                    <label>Application Status</label>
                    <input type="text" name="application_status" id="application_status" value='<%Response.write rs("application_status") %>' readonly=true/>
					
                    <label style="width:150px">Admission Decision</label>
                    <input type="text" name="admission_decision" id="admission_decision" value='<%Response.write rs("admission_decision") %>' readonly=true/>
                    
                    &nbsp 
                    <label>Degree Program</label>
					<input type="text" name="Degree_Program" id="Degree_Program" value='<%Response.write rs("Degree_Program") %>' readonly=true />
                   
                    
                    
                    <br/><br/><br/><br/>
                    
                     <label>Date of Initial Entry</label>
					<input type="text" name="DateofInitialEntry" class="date"  id="DateofInitialEntry" value='<%Response.write rs("DateofInitialEntry") %>' readonly=true/>
                    <label style="width:160px">OAR Application Date</label>
					<input type="text" name="oar_application_date" class="date"  id="oar_application_date" value='<%Response.write rs("oar_application_date") %>' readonly=true/>
                    <label>Reapplicant</label>
                    <input type="checkbox" disabled="disabled" style="width:20px;height:20px;" name="Reapplicant" id="Reapplicant" class="checkboxField" value='<%Response.write rs("Reapplicant") %>'/>
                   
                    
                    <br /><br /><br /><br />   
                    <label>Decision Date</label>
                    <input type="text" name="decision_dt" id="decision_dt" class="date" value='<%Response.write rs("decision_dt") %>' readonly=true />
                    <label style = "width:180px">Decision Letter Sent Date</label>
                    <input type="text" name="Decision_Letter_Sent_Date" id="Decision_Letter_Sent_Date" class="date" value='<%Response.write rs("Decision_Letter_Sent_Date") %>' readonly=true/>
                    
                    <br /><br /><br /><br />    
                    <label>Confirmed</label>        
                    <input type="checkbox" disabled="disabled" style="width:20px;height:20px;" name="confirmed" id="confirmed" class="checkboxField" value='<%Response.write rs("confirmed") %>'/>
                    <label>Confirmed Date</label>
                    <input type="text" name="Confirmed_Dt" id="Confirmed_Dt" class="date" value='<%Response.write rs("Confirmed_Dt") %>' readonly=true/>
                    <label>Admit Term</label>
                    <input type="text" name="Admit_Term" id="Admit_Term" value='<%Response.write rs("Admit_Term") %>' readonly=true/>

                    


                    <br /><br /><br /><br />
                    <label style="width:180px">Financial Aid Requested</label>
                    <input type="checkbox" disabled="disabled" style="width:20px;height:20px;" name="Financial_Aid_Requested" id="Financial_Aid_Requested" class="checkboxField" value='<%Response.write rs("Financial_Aid_Requested") %>'/>
                    <label>UIC Employee</label>
                    <input type="checkbox" disabled="disabled" style="width:20px;height:20px;" name="UIC_employee" id="UIC_employee" class="checkboxField" value='<%Response.write rs("UIC_employee") %>'/>
                    <label>Orientation</label>
                    <input type="checkbox" disabled="disabled" style="width:20px;height:20px;" name="Orientation" id="Orientation" class="checkboxField" value='<%Response.write rs("Orientation") %>'/>
                    
                    <label style="width:150px">Information Session</label>
                    <input type="checkbox" disabled="disabled" style="width:20px;height:20px;" name="Open_house" id="Open_house" class="checkboxField" value='<%Response.write rs("Open_house") %>'/>
                    
                    <br /><br /><br /><br /><br />
                    <label>UIC UG/GRAD apps</label>
                    <input type="checkbox" disabled="disabled" style="width:20px;height:20px;" name="UIC_UG_Grad_Apps" id="UIC_UG_Grad_Apps" class="checkboxField" value='<%Response.write rs("UIC_UG_Grad_Apps") %>' />
                     
                    <label>Application Fee</label>
                    <input type="text" name="Application_Fee" id="Application_Fee" value='<%Response.write rs("Application_Fee") %>' readonly=true/>
                    
                    <label style="width:180px">Jane Addams Application</label>
                    <input type="checkbox" disabled="disabled" style="width:20px;height:20px;" name="Jane_Addams_appln" id="Jane_Addams_appln" class="checkboxField" value='<%Response.write rs("Jane_Addams_appln") %>'/>
                    <br /><br /><br /><br />
                    <label>Transcripts</label>
                    <input type="checkbox" disabled="disabled" style="width:20px;height:20px;" name="Transcripts" id="Transcripts" class="checkboxField" value='<%Response.write rs("Transcripts") %>'/>
                    <label>TOEFL Score</label>
					<input type="text" name="TOEFL_Score" id="TOEFL_Score" value='<%Response.write rs("TOEFL_Score") %>' readonly=true/>
                    <br /><br /><br /><br />
                    <label>GRE Quantitative</label>
					<input type="text" name="GRE_Quantitative" id="GRE_Quantitative" value='<%Response.write rs("GRE_Quantitative") %>' readonly=true />
                    <label>GRE Verbal</label>
					<input type="text" name="GRE_Verbal" id="GRE_Verbal" value='<%Response.write rs("GRE_Verbal") %>' readonly=true />
                    <label>GRE Analytical</label>
					<input type="text" name="GRE_Analytical" id="GRE_Analytical" value='<%Response.write rs("GRE_Analytical") %>' readonly=true />

                    <br /><br /><br /><br />
                    <label>Field of Interest</label>
					<input type="text" name="Field_of_Interest" id="Field_of_Interest" value='<%Response.write rs("Field_of_Interest") %>' readonly=true/>
                    <label style="width:190px">Dec and Cert of Finances Sub</label>
                    <input type="checkbox" disabled="disabled" style="width:20px;height:20px;" name="Dec_Cert_Finances_Sub" id="Dec_Cert_Finances_Sub" class="checkboxField" value='<%Response.write rs("Dec_Cert_Finances_Sub") %>'/>
                    
                    <br/><br/><br/><br/>
                    </p>
                    </fieldset>
					
                    <div id="application2" align="center">
                    <fieldset class="step">
                    <legend></legend>
                    <p>
                    <br/>
                    <label>Entered By</label>
                    <input type="text" name="enteredby" id="enteredby" value='<%Response.write rs("enteredby") %>' readonly=true/>
                    <label>Last Updated Date</label>
                    <input type="text" name="LastUpdatedDt" id="LastUpdatedDt" value='<%Response.write rs("LastUpdatedDt") %>' readonly=true/>
                    
                    <br /><br /><br /><br />
                    <label>Comments</label>
                    <textarea id="comments" name="comments" cols="70" rows="5" readonly=true><%Response.write rs("Comments") %></textarea>
                     <br /><br /><br /><br />
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
