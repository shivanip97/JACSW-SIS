<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | Add New Student</title>
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
            <div id="steps">
				<form id="studentForm" method="post" action="AfterAddStudent.asp">
					<h3>Add New Student</h3>
                     <br/>
                    <a href=ShowStudents.asp>Back to Show Students</a> 
                    <br/> <br/>
                    <p>
                    <label>Add Student Form</label>
                    <br/><br/><br/>
                    <label>First Name</label>
					<input type="text" name="firstname" required id="firstname"/>   
                    <label>Middle Name</label>
					<input type="text" name="middlename" id="middlename"/> 
                    <label>Last Name</label>
					<input type="text" name="lastname" required id="lastname"/>    
                    <br/><br/><br/>
                    <label>Banner ID</label>
					<input type="text" name="uin" required id="uin"/> 
                    <label>Date of Birth</label>
					<input type="text" name="dob" class="date" required id="dob"/> 
                    <label>Maiden Name</label>
					<input type="text" name="maidenname" required id="maidenname" />
                    <label>OAR Application Date</label>
					<input type="text" name="appdate" class="date" required id="appdate" />
                    <br/><br/><br/>
                    <label>Gender</label>
   	                <select name="gender" id="gender">
         			<option value="0">-- Select --</option>
  					<option value="Male">Male</option>
					<option value="Female">Female</option>
				    </select>
                    <br/><br/><br/>
                    <label>Current Address 1</label>
					<input type="text" name="currentAddress1" required id="currentAddress1" />
                    <label>Current Address 2</label>
					<input type="text" name="currentAddress2" required id="currentAddress2" />
                    <label>Current City</label>
					<input type="text" name="currentcity" required id="currentcity" />
                    <br/><br/><br/>
                    <label>Current State</label>
                    <select name="currentstate" id="currentstate">
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
                    <label>Current Zip</label>
					<input type="text" name="currentzip" class="zip" required id="currentzip" />
                    <label>Current Country</label>
					<input type="text" name="currentcountry" required id="currentcountry" />
                    <br/><br/><br/><br/>
                    <label>Home Phone</label>
					<input type="text" name="homephone" class="homephone" required id="homephone" />
                    <label>Work Phone</label>
					<input type="text" name="workphone" class="workphone" required id="workphone" />
                    <label>International Phone</label>
					<input type="text" name="intphone" class="iphone" id="intphone" />
                    <br/><br/><br/><br/>
                     <label>UG College</label>
					<input type="text" name="ugcollege" id="ugcollege" />
                    <label>UG GPA</label>
					<input type="text" name="uggpa" class="gpa" id="uggpa" />
                    <label>UG Major</label>
					<input type="text" name="ugmajor" id="ugmajor" />
                    <br/><br/><br/><br/>
                     <label>Grad College</label>
					<input type="text" name="gradcollege" id="gradcollege" />
                    <label>Grad GPA</label>
					<input type="text" name="gradgpa" class="gpa" id="gradgpa" />
                    <label>Grad Major</label>
					<input type="text" name="gradmajor" id="gradmajor" />
                    <br/><br/><br/><br/>
                    <label>Grad Degree</label>
					<input type="text" name="graddegree" id="graddegree" />
                    <label>Email</label>
					<input type="text" name="email" required id="email" />
                    <br/><br/><br/><br/>
                    <label>Comments</label>
                    <textarea id="comments" name="comments" cols="70" rows="10" ></textarea>
                    <br/><br/><br/>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/><br/><br/>
					<button type="submit" name="Submit" onclick="return validate();">Add Student</button>
					<br/><br/>
                    </p>
				</form>
               </div>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</body>
</html>
