<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>SIS | View Student Info</title>
    <!--#include file="DBconn.asp"-->
    <!--#include file="Login_Check.asp"-->
    <!--#include file="header.asp"-->
    <% 
        ErrMsg = Request("ErrMsg")
        UIN = Request("Button1")
    %>
    <link rel="stylesheet" href="css/tabstyle.css" type="text/css" media="screen" />
	<script type="text/javascript" src="jquery/jquery-1.8.2.min.js"></script>
    <script type="text/javascript" src="jquery/jquery-ui-1.8.2.custom.min.js"></script>
    <script type="text/javascript" src="jquery/sliding.form.js"></script>

    <script>
        $(document).ready(function () {
        });
    </script>

        <style type="text/css">
		table {
			text-align: left;
			font-size: 12px;
			font-family: verdana;
			background: #c0c0c0;
			table-layout:fixed;
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
			height:50px;
			overflow: hidden;
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

<body>
    <div id="content">
        <h1>SIS - Student Information</h1>
        <br /><br />
          <a style="font-size:12pt;" href="ShowStudents.asp"> HOME</a>
        <br /><br />
        <br />
        <div id="wrapper">
        <div id="navigation" style="display: none;">
                <ul>
                    <li><a href="#">Applicant</a></li>
                    <li><a href="#">Current</a></li>
                     <li><a href="#">Field</a></li>
                </ul>
            </div>
            <div id="steps" align="center">

            <form id="formElem" name="formElem" action="" method="post">
                <fieldset class="step">
                <legend></legend>
                <% 
                    set rs=Server.CreateObject("ADODB.recordset")
                    query="select * from Applicants where UIN ='"& UIN &"'"
                    rs.Open query,conn
                %>
                <p>
                    <br/>
                    <label>First Name</label>
					<input type="text" name="firstname"  id="firstname" value='<%Response.write rs("firstname") %>' readonly=true/>   
                    <label>Middle Name</label>
					<input type="text" name="middlename" id="middlename" value='<%Response.write rs("middlename") %>' readonly=true/> 
                    <label>Last Name</label>
					<input type="text" name="lastname"  id="lastname" value='<%Response.write rs("lastname") %>' readonly=true/>    
                    <br/><br/><br/><br/>
                    <label>Maiden Name</label>
					<input type="text" name="maidenname"  id="maidenname" value='<%Response.write rs("maidenname") %>' readonly=true/>
                    <label>UIN</label>
					<input type="text" name="uin"  id="uin" value='<%Response.write rs("uin") %>' readonly=true/> 
                    <label>Date of Birth</label>
					<input type="text" name="dob" class="date"  id="dob" value='<%Response.write rs("DateOfBirth") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    <label>Gender</label>
					<input type="text" name="gender"  id="gender" value='<%Response.write rs("gender") %>' readonly=true/>
                    <label>OAR Application Date</label>
					<input type="text" name="oar_application_date" class="date"  id="oar_application_date" value='<%Response.write rs("oar_application_date") %>' readonly=true/>  
                    <label>Application Status</label>
					<input type="text" name="application_status"  id="application_status" value='<%Response.write rs("application_status") %>' readonly=true/>                  
                    <br/><br/><br/><br/>
                    <label>Ready for Review Date</label>
					<input type="text" name="readyforreviewdate" class="date"  id="readyforreviewdate" value='<%Response.write rs("readyforreviewdate") %>' readonly=true/>  
                    <label>Entered By</label>
					<input type="text" name="enteredby"  id="enteredby" value='<%Response.write rs("enteredby") %>' readonly=true/> 
                    <label>SSN</label>
					<input type="text" name="ssn"  id="ssn" value='<%Response.write rs("ssn") %>' readonly=true/> 
                    <br/><br/><br/><br/>
                    <label>Current Address 1</label>
					<input type="text" name="currentAddress1"  id="currentAddress1" value='<%Response.write rs("currentAddress1") %>' readonly=true/>
                    <label>Current Address 2</label>
					<input type="text" name="currentAddress2"  id="currentAddress2" value='<%Response.write rs("currentAddress2") %>' readonly=true/>
                    <label>Current City</label>
					<input type="text" name="currentcity"  id="currentcity" value='<%Response.write rs("currentcity") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Current ZipCode</label>
					<input type="text" name="currentzipcode" class="zip" required id="currentzipcode" value='<%Response.write rs("currentzipcode") %>' readonly=true/>
                    <label>Current State</label>
					<input type="text" name="currentstate"  id="currentstate" value='<%Response.write rs("currentstate") %>' readonly=true/>
                    <label>Current Country</label>
					<input type="text" name="currentcountry"  id="currentcountry" value='<%Response.write rs("currentcountry") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Mailing Address</label>
					<input type="text" name="mailingaddress"  id="mailingaddress" value='<%Response.write rs("mailingaddress") %>' readonly=true/>
                    <label>Home Phone</label>
					<input type="text" name="homephone" class="homephone" required id="homephone" value='<%Response.write rs("homephone") %>' readonly=true/>
                    <label>Work Phone</label>
					<input type="text" name="workphone" class="workphone" required id="workphone" value='<%Response.write rs("workphone") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>International Phone</label>
					<input type="text" name="internationalphonenumber" class="iphone" id="internationalphonenumber" value='<%Response.write rs("InternationalPhoneNumber") %>' readonly=true/>
                    <label>Email</label>
					<input type="text" name="email" id="email" value='<%Response.write rs("email") %>' readonly=true/>
                    <label>Degree Program</label>
					<input type="text" name="degreeprogram" id="degreeprogram" value='<%Response.write rs("Degree_Program") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Credit in BA BS</label>
					<input type="text" name="credit_in_ba_bs" id="credit_in_ba_bs" value='<%Response.write rs("credit_in_ba_bs") %>' readonly=true/>
                    <label>Credit in English</label>
					<input type="text" name="credit_in_english" id="credit_in_english" value='<%Response.write rs("credit_in_english") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Requesting Schools</label>
					<input type="text" name="requesting_schools" id="requesting_schools" value='<%Response.write rs("requesting_schools") %>' readonly=true/>
                    
                    <label>Reapplicant</label>
					<input type="text" name="Reapplicant" id="Reapplicant" value='<%Response.write rs("Reapplicant") %>' readonly=true/>
                    
                    <label>Program Type</label>
					<input type="text" name="program_type" id="program_type" value='<%Response.write rs("program_type") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Concentration</label>
					<input type="text" name="concentration" id="concentration" value='<%Response.write rs("concentration") %>' readonly=true/>
                    
                    <label>Admission Decision</label>
					<input type="text" name="admission_decision" id="admission_decision" value='<%Response.write rs("admission_decision") %>' readonly=true/>
                    
                    <label>Decision Date</label>
					<input type="text" name="decision_dt" id="decision_dt" class="date" value='<%Response.write rs("decision_dt") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Decision Letter Sent Date</label>
					<input type="text" name="Decision_Letter_Sent_Date" id="Decision_Letter_Sent_Date" class="date" value='<%Response.write rs("Decision_Letter_Sent_Date") %>' readonly=true/>
                    
                    <label>Limited Status</label>
					<input type="text" name="Limited_status" id="Limited_status" value='<%Response.write rs("Limited_status") %>' readonly=true/>
                    
                    <label>Confirmed</label>
					<input type="text" name="Confirmed" id="Confirmed" value='<%Response.write rs("Confirmed") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Confirmed Date</label>
					<input type="text" name="Confirmed_Dt" id="Confirmed_Dt" class="date" value='<%Response.write rs("Confirmed_Dt") %>' readonly=true/>
                    
                    <label>Admit Term</label>
					<input type="text" name="Admit_Term" id="Admit_Term" value='<%Response.write rs("Admit_Term") %>' readonly=true/>
                    
                    <label>Credit In Statistics</label>
					<input type="text" name="Credit_in_Statistics" id="Credit_in_Statistics" value='<%Response.write rs("Credit_in_Statistics") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Financial Aid Request</label>
					<input type="text" name="Credit_in_Statistics" id="Credit_in_Statistics" value='<%Response.write rs("Credit_in_Statistics") %>' readonly=true/>
                    
                    <label>Credit In Statistics</label>
					<input type="text" name="Credit_in_Statistics" id="Credit_in_Statistics" value='<%Response.write rs("Credit_in_Statistics") %>' readonly=true/>
                    
                    <label>Financial Aid Request</label>
					<input type="text" name="Financial_Aid_Request" id="Financial_Aid_Request" value='<%Response.write rs("Financial_Aid_Request") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Basic Skill Test</label>
					<input type="text" name="Basic_Skill_Test" id="Basic_Skill_Test" value='<%Response.write rs("Basic_Skill_Test") %>' readonly=true/>
                    <label>Passed Test</label>
					<input type="text" name="Passed_Test" id="Passed_Test" value='<%Response.write rs("Passed_Test") %>' readonly=true/>
                    <label>UG College</label>
					<input type="text" name="ugcollege" id="ugcollege" value='<%Response.write rs("UGCollege") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>UG GPA</label>
					<input type="text" name="uggpa" class="gpa" id="uggpa" value='<%Response.write rs("UGGPA") %>' readonly=true/>
                    <label>UG Major</label>
					<input type="text" name="ugmajor" id="ugmajor" value='<%Response.write rs("UGMajor") %>' readonly=true/>
                    
                    <label>Grad College</label>
					<input type="text" name="gradcollege" id="gradcollege" value='<%Response.write rs("GradCollege") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Grad GPA</label>
					<input type="text" name="gradgpa" class="gpa" id="gradgpa" value='<%Response.write rs("GradGPA") %>' readonly=true/>
                    <label>Grad Major</label>
					<input type="text" name="gradmajor" id="gradmajor" value='<%Response.write rs("GradMajor") %>' readonly=true/>
                    
                    <label>Grad Degree</label>
					<input type="text" name="graddegree" id="graddegree" value='<%Response.write rs("GradDegree") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Comments</label>
                    <textarea id="comments" name="comments" cols="70" rows="5" readonly=true><%Response.write rs("Comments") %></textarea>
                    <label>Date Comments Entered</label>
                    <textarea id="DateCommentsEntered" name="DateCommentsEntered" cols="30" rows="3" readonly=true><%Response.write rs("DateCommentsEntered") %></textarea>
                    <label>Last Updated Date</label>
					<input type="text" name="LastUpdatedDt" id="LastUpdatedDt" class="date" value='<%Response.write rs("LastUpdatedDt") %>' readonly=true/>
                    
                    <br/><br/><br/><br/>
                    </p>
                    
                 </fieldset>
            <div id="current" align="center">

                <fieldset class="step">
                <legend></legend>
                <% 
                    set rs=Server.CreateObject("ADODB.recordset")
                    query="select * from CurrentStudents where UIN ='"& UIN &"'"
                    rs.Open query,conn
                    
                %>
                <p>
                    <br/>
                  <label>First Name</label>
					<input type="text" name="firstname"  id="fstname" value='<%Response.write rs("firstname") %>' readonly=true/>   
                    <label>Middle Name</label>
					<input type="text" name="middlename" id="mname" value='<%Response.write rs("middlename") %>' readonly="true"/> 
                    <label>Last Name</label>
					<input type="text" name="lastname"  id="lname" value='<%Response.write rs("lastname") %>' readonly=true/>    
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
                    <label>Maiden Name</label>
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
                    
                    <label>Email</label>
					<input type="text" name="email" id="email" value='<%Response.write rs("email") %>' readonly=true/>
                    <label>Limited Status</label>
					<input type="text" name="LimitedStatus" required id="LimitedStatus" value='<%Response.write rs("LimitedStatus") %>' readonly=true/>
					<br/><br/><br/><br/>
                    <label>Comments</label>
                    <textarea id="comments" name="comments" cols="45" rows="5" readonly=true><%Response.write rs("Comments") %></textarea>
                    <br/><br/><br/><br/>
                    
                    <label>Program Type</label>
					<input type="text" name="ProgramType" required id="ProgramType" value='<%Response.write rs("ProgramType") %>' readonly=true/>  
                    <label>Concentration</label>
					<input type="text" name="Concentration" required id="Concentration" value='<%Response.write rs("Concentration") %>' readonly=true/>                  
                    <label>Decision</label>
					<input type="text" name="Decision" required id="Decision" value='<%Response.write rs("Decision") %>' readonly=true/>  
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
                    <label>Current Year</label>
					<input type="text" name="CurrentYear" required id="CurrentYear" value='<%Response.write rs("CurrentYear") %>' readonly=true/>
                    <br/><br/><br/><br/>

                    <label>Applying For Graduation</label>
					<input type="text" name="ApplyingForGraduation" required id="ApplyingForGraduation" value='<%Response.write rs("ApplyingForGraduation") %>' readonly=true/>
                    <label>Graduation Term Applied For</label>
					<input type="text" name="GraduationTermAppliedFor" required id="GraduationTermAppliedFor" value='<%Response.write rs("GraduationTermAppliedFor") %>' readonly=true/>
                    <label>Term Graduated</label>
					<input type="text" name="TermGraduated" required id="TermGraduated" value='<%Response.write rs("TermGraduated") %>' readonly=true/>
                    <br/><br/><br/><br/>

                    <label>Degree Applying For</label>
					<input type="text" name="DegreeApplyingFor" required id="DegreeApplyingFor" value='<%Response.write rs("DegreeApplyingFor") %>' readonly=true/>
                    <label>Mailbox Number</label>
					<input type="text" name="MailboxNumber" required id="MailboxNumber" value='<%Response.write rs("MailboxNumber") %>' readonly=true/>
                              
                    <br/><br/><br/><br/>
                    </p>
                    
                 </fieldset>
            <div id="field" align="center">

                <fieldset class="step">
                <legend></legend>
                <% 
                    set rs=Server.CreateObject("ADODB.recordset")
                    query="select * from Field where UIN ='"& UIN &"'"
                    rs.Open query,conn
                %>
                <p>
                    <br/>
					<label>Field Type</label>
					<input type="text" name="fieldtype" required id="fieldtype" value='<%Response.write rs("FieldType") %>' readonly=true/>   
                    <label>POE</label>
					<input type="text" name="poe" id="poe" value='<%Response.write rs("POE") %>'readonly=true/> 
                    <label>Banner</label>
					<input type="text" name="uin" required id="uin" value='<%Response.write rs("UIN") %>'readonly=true/>    
                    <br/><br/><br/><br/>
					
                    <label>Faculty Liasion Foundation</label>
					<input type="text" name="FacultyLiasionFoundation" required id="FacultyLiasionFoundation" value='<%Response.write rs("FacultyLiasionFoundation") %>' readonly=true/>   
                    <label>Faculty Liasion Concentration</label>
					<input type="text" name="FacultyLiasionConcentration" required id="FacultyLiasionConcentration" value='<%Response.write rs("FacultyLiasionConcentration") %>' readonly=true/>   
					<label>Working Liasion Concentration</label>
					<input type="text" name="WorkingLiasionConcentration" required id="WorkingLiasionConcentration" value='<%Response.write rs("WorkingLiasionConcentration") %>' readonly=true/>   
                    <br/><br/><br/><br/>
					
                    <label>Working Liasion Foundation</label>
					<input type="text" name="WorkingLiasionFoundation" required id="WorkingLiasionFoundation" value='<%Response.write rs("WorkingLiasionFoundation") %>' readonly=true/>   
                    <label>Info Sent</label>
                    <input type="text" name="infoSent" class="date" required id="infoSent" value='<%Response.write rs("InfoSent") %>' readonly=true/> 
                    <label>Working Liasion Concentration Term</label>
					<input type="text" name="WorkingLiasionConcentrationTerm" required id="WorkingLiasionConcentrationTerm" value='<%Response.write rs("WorkingLiasionConcentrationTerm") %>' readonly=true/>   
                    <br/><br/><br/><br/><br />
					
                    <label>Working Liasion Foundation Term</label>
					<input type="text" name="WorkingLiasionFoundationTerm" required id="WorkingLiasionFoundationTerm" value='<%Response.write rs("WorkingLiasionFoundationTerm") %>' readonly=true/>   
                    <label>Date comments entered</label>
					<input type="text" name="dce" class="date" required id="dce" value='<%Response.write rs("DateCommentsEntered") %>' readonly=true/>
                    <br/><br/><br/><br/>
                    <label>Comments</label>
					<textarea id="comments" name="comments" cols="70" rows="5" readonly=true><%Response.write rs("Comments") %></textarea>
   
                    <br/><br/><br/><br/>
                    </p>
                    
                 </fieldset>

                </form>
                </div>
                 <% %> 
            </div>
        </div>
    <!--#include file="footer.asp"-->
</body>
</html>

