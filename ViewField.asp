<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UIN = Request("Button1")
set rs=Server.CreateObject("ADODB.recordset")
query="select * from Field1 where UIN ='"& UIN &"'"
rs.Open query,conn
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<!--#include file="header.asp"-->
<title>SIS | Edit Field</title>
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
        $('#fieldForm').find(':input:not(button)').each(function () {
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
            alert('Please complete the form by filling in fields highlighted in Red.')
        }
        return shouldProceed;
    }
 	</script>
</head>
<body>
    <div id="content" align=center>
            <div id="steps">
				<form id="fieldForm" method="post" action="AfterUpdateField.asp">
					<h3>Edit Field Information</h3>
                     <br/>
                    <a href="ViewFieldNew.asp?UIN=<%Response.Write (UIN) %>">Back to Show Field</a> 
                    <br/> <br/>
                    <p>
                    <label>Field Form</label>
                    <br/><br/><br/>
                    
                    <label>UIN</label>
					<input type="text" name="uin" readonly ="true" required id="uin" value='<%Response.write rs("UIN") %>'/>
                        <%
                    set rs1=Server.CreateObject("ADODB.recordset")
					course_query3="select FirstName, LastName,ProgramType from CurrentStudents where UIN = '"& UIN & "'"
					rs1.Open course_query3,conn
                    If not rs1.EOF Then
                    %>
                        <label>Program Type</label>
					<input type="text" name="programtype" readonly ="true" required id="programtype" readonly ="true" value='<%Response.write rs1("ProgramType") %>'/>   
                    
                            <label>Name</label>
					<input type="text" name="name" required id="name" readonly ="true" value='<% Response.Write rs1("FirstName") %> &nbsp <% Response.Write rs1("LastName") %>'/>
            
                  <% End If %>     
                   
                     <br />   <br/><br/> <br/><br />

                    <label>Working Liasion Generalist</label>
   	                <select name="wlf" id="wlf">
         			<option value="<% Response.write rs("WorkingLiasionFoundation") %>"><% Response.write rs("WorkingLiasionFoundation") %></option>
  					<option value="">Select</option>
                    <option value="B. Coats">B. Coats</option>
                    <option value="N. Rosales">N. Rosales</option>
                    <option value="A. Johnson">A. Johnson</option>
                    <option value="K. Jenkins">K. Jenkins</option>
                    <option value="C. Melka">C. Melka</option>
                    <option value="C. Taylor">C. Taylor</option>
                    <option value="J. Fisher">J. Fisher</option>
                    <option value="M. Carrasco">M. Carrasco</option>
                    </select>

                        
                   
                        <label>Field Liasion Generalist</label>
					<select name="flf" id="flf">
                    <option value="<% Response.write rs("FacultyLiasionFoundation") %>"><% Response.write rs("FacultyLiasionFoundation") %></option>
                    <option value="">Select</option>
                    <option value="B. Coats">B. Coats</option>
                    <option value="N. Rosales">N. Rosales</option>
                    <option value="A. Johnson">A. Johnson</option>
                    <option value="K. Jenkins">K. Jenkins</option>
                    <option value="C. Melka">C. Melka</option>
                    <option value="C. Taylor">C. Taylor</option>
                    <option value="J. Fisher">J. Fisher</option>
                    <option value="M. Carrasco">M. Carrasco</option>
				    </select>

                         <label>Generalist Term</label>
   	                <select name="wlft" id="wlft">
         			<option value="<% Response.write rs("WorkingLiasionFoundationTerm") %>"><% Response.write rs("WorkingLiasionFoundationTerm") %></option>
                    <option value="">Select</option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Summer 2015">Summer 2015</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Summer 2016">Summer 2016</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
					<option value="Summer 2017">Summer 2017</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Summer 2018">Summer 2018</option>
                    <option value="Fall 2018">Fall 2018</option>
                    <option value="Spring 2019">Spring 2019</option>
                    <option value="Summer 2019">Summer 2019</option>
                    <option value="Fall 2019">Fall 2019</option>
                    <option value="Spring 2020">Spring 2020</option>
                    <option value="Summer 2020">Summer 2020</option>
                    <option value="Fall 2020">Fall 2020</option>
                    <option value="Spring 2021">Spring 2021</option>
                    <option value="Summer 2021">Summer 2021</option>
                    <option value="Fall 2021">Fall 2021</option>
                    <option value="Spring 2022">Spring 2022</option>
                    <option value="Summer 2022">Summer 2022</option>
                    <option value="Fall 2022">Fall 2022</option>
                    <option value="Spring 2023">Spring 2023</option>
                    <option value="Summer 2023">Summer 2023</option>
                    <option value="Fall 2023">Fall 2023</option>
				    </select>
                        <br />  <br/><br/>  <br/><br/><br/>
                    
                    <label>Working Liasion Specialization</label>
   	                <select name="wlc" id="wlc">
                    <option value="<% Response.write rs("WorkingLiasionConcentration") %>"><% Response.write rs("WorkingLiasionConcentration") %></option>
                    <option value="">Select</option>
                    <option value="B. Coats">B. Coats</option>
                    <option value="N. Rosales">N. Rosales</option>
                    <option value="A. Johnson">A. Johnson</option>
                    <option value="K. Jenkins">K. Jenkins</option>
                    <option value="C. Melka">C. Melka</option>
                    <option value="C. Taylor">C. Taylor</option>
                    <option value="J. Fisher">J. Fisher</option>
                    <option value="M. Carrasco">M. Carrasco</option>
				    </select>
                 
                    
                    
                        <label>Field Liasion Specialization</label>
					<select name="flc" id="flc">
                    <option value="<% Response.write rs("FacultyLiasionConcentration") %>"><% Response.write rs("FacultyLiasionConcentration") %></option>
                    <option value="">Select</option>
                    <option value="B. Coats">B. Coats</option>
                    <option value="N. Rosales">N. Rosales</option>
                    <option value="A. Johnson">A. Johnson</option>
                    <option value="K. Jenkins">K. Jenkins</option>
                    <option value="C. Melka">C. Melka</option>
                    <option value="C. Taylor">C. Taylor</option>
                    <option value="J. Fisher">J. Fisher</option>
                    <option value="M. Carrasco">M. Carrasco</option>
				    </select>

                        <label>Specialization Term</label>
   	                <select name="wlct" id="wlct">
         			<option value="<% Response.write rs("WorkingLiasionConcentrationTerm") %>"><% Response.write rs("WorkingLiasionConcentrationTerm") %></option>
  					<option value="">Select</option>
                    <option value="Spring 2015">Spring 2015</option>
                    <option value="Summer 2015">Summer 2015</option>
                    <option value="Fall 2015">Fall 2015</option>
                    <option value="Spring 2016">Spring 2016</option>
                    <option value="Summer 2016">Summer 2016</option>
                    <option value="Fall 2016">Fall 2016</option>
                    <option value="Spring 2017">Spring 2017</option>
					<option value="Summer 2017">Summer 2017</option>
                    <option value="Fall 2017">Fall 2017</option>
                    <option value="Spring 2018">Spring 2018</option>
                    <option value="Summer 2018">Summer 2018</option>
                    <option value="Fall 2018">Fall 2018</option>
                    <option value="Spring 2019">Spring 2019</option>
                    <option value="Summer 2019">Summer 2019</option>
                    <option value="Fall 2019">Fall 2019</option>
                    <option value="Spring 2020">Spring 2020</option>
                    <option value="Summer 2020">Summer 2020</option>
                    <option value="Fall 2020">Fall 2020</option>
                    <option value="Spring 2021">Spring 2021</option>
                    <option value="Summer 2021">Summer 2021</option>
                    <option value="Fall 2021">Fall 2021</option>
                    <option value="Spring 2022">Spring 2022</option>
                    <option value="Summer 2022">Summer 2022</option>
                    <option value="Fall 2022">Fall 2022</option>
                    <option value="Spring 2023">Spring 2023</option>
                    <option value="Summer 2023">Summer 2023</option>
                    <option value="Fall 2023">Fall 2023</option>
				    </select>
                    <br/><br/><br/><br/><br/><br/>

                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/><br/><br/>
                    <button type="submit" name="Submit1" onclick="this.form.action='AfterUpdateField.asp?UIN=' + this.value; this.forms.submit();" value='<%Response.Write (UIN) %>'>Save Field</button><br /><br />
					<button type="submit" name="Submit" onclick="this.form.action='RemoveFieldConfirm.asp?UIN=' + this.value; this.forms.submit();" value='<%Response.Write (UIN) %>'>Remove Field</button><br /><br />
					<br/><br/>
                    </p>
				</form>
               </div>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</body>
</html>
