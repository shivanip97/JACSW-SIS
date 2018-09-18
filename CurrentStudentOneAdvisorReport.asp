<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Login_Check.asp"-->
<!--#include file="fpdf.asp"-->
<!--#include file="DBconn.asp"-->

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
	        $('.editableRace').editable('UpdateStudent.asp', {
	            data: " {'Did Not Answer':'Did Not Answer','Native American/Alaskan Native':'Native American/Alaskan Native','African/AfricanAmerican':'African/AfricanAmerican','Asian/Pacific Islander':'Asian/Pacific Islander','Caucasian':'Caucasian','Hispanic':'Hispanic','International':'International','selected':'Did Not Answer'}",
	            type: 'select',
	            submit: 'OK'
	        });
	        var adterm = $("#admit_term option:selected").text();

	        document.getElementById("termname").innerHTML = adterm;
	    });

	    function getval(sel) {
	        window.location = "https://socialwork.cc.uic.edu/SIS/CurrentStudentOneAdvisorFinalReport.asp?ID=" + sel.value;
	    }
 	</script>
<h4><a href="ShowCurrentStudents.asp">Home</a> |<a href="MSWCurrentReports.asp">Current Student Reports</a>  | <a href="logout.asp">Log Out</a></h4>
<div align="center"> 
 <label style="font-size: 1.17em;font-weight: bold;">Select Advisor</label>
      <select name="Advisor" id="Advisor" onchange="getval(this);">
                    <option value="All">All</option>
                    <option value="Aaron Gottlieb">Aaron Gottlieb</option>
                    <option value="Bonecutter">Bonecutter</option>
                    <option value="Branden McLeod">Branden McLeod</option>
                    <option value="Butterfield">Butterfield</option>         
                    <option value="Doyle">Doyle</option>
                    <option value="Fisher">Fisher</option>
                    <option value="Gaston">Gaston</option>
                    <option value="Geiger">Geiger</option>
                    <option value="Gleeson">Gleeson</option>
                    <option value="Hairston">Hairston</option>
                    <option value="Hounmenou">Hounmenou</option>
                    <option value="Hsieh">Hsieh</option>
                    <option value="Jack Lu">Jack Lu</option>
                    <option value="Johnson">Johnson</option>
                    <option value="Karen D'Angelo">Karen D'Angelo</option> 
                    <option value="Leathers">Leathers</option>
                    
                    <option value="McCoy">McCoy</option>
                    <option value="McKay-Jackson">McKay-Jackson</option>
                    <option value="Mitchell">Mitchell</option>
                    
                    <option value="O'Brien">O'Brien</option>
                    <option value="Robert Wilson">Robert Wilson</option>
                    <option value="Swartz">Swartz</option>
                    <option value="Watson">Watson</option>
                    </select> 
      <br /><br />
    </div>


    