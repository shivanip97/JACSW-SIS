﻿<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UIN = Request("UIN")
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
<body bgcolor="#f2f2f2">

<!--#include file="headerField.asp"-->
    <div align="center">
     
                <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong><br/>
     
           
     <p><label for="UIN">UIN: </label><strong><font color="#000000"><% Response.Write(UIN) %> </font></strong></p>

     
        <%
                    set rs=Server.CreateObject("ADODB.recordset")
					course_query1="select FirstName, LastName from CurrentStudents where UIN = '"& UIN & "'"
					rs.Open course_query1,conn
                    If rs.EOF Then
                   Else
            %>
                    <p><label for="Name:">Name: </label><strong><font color="#000000"><% Response.Write rs("FirstName") %> &nbsp <% Response.Write rs("LastName") %> </font></strong></p>
                    
                  <% End If %>
     
        <%
                    set rs=Server.CreateObject("ADODB.recordset")
					course_query="select * from Field1 where UIN like '"& UIN & "'"
					rs.Open course_query,conn
                    %>  
            
  <div id="content" align=center>
            <div id="steps">
				<form id="fieldForm" method="post" action="">        
					<h3>View Field Information</h3>
                     <br/>
                    <a href="ViewFieldNew.asp?UIN=<%Response.Write (UIN) %>">Back to Show Field</a> 
                    
                    <br/> <br/>
                    <p>
                    <label>Field Form</label>
                    <br/><br/><br/>
                    
                    <label>UIN</label>
					<input type="text" name="uin" required id="uin" readonly ="true" value='<%Response.write rs("UIN") %>'/>   
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
                    <br/><br/><br/>
                         <label>Field Type Year</label>
                    <input type="text" name="fieldtypeyear" required id="fieldtypeyear" readonly="true" value='<%Response.write rs("FieldTypeYear")  %>' />
                        
					<br/><br/><br/>
                    <label>Working Liasion Foundation</label>
                    <input type="text" name="wlf" id="wlf" readonly ="true" value='<%Response.write rs("WorkingLiasionFoundation") %>'/> 
                        <label>Faculty Liasion Foundation</label>
                    <input type="text" name="flf" id="flf" readonly ="true" value='<%Response.write rs("FacultyLiasionFoundation") %>'/>
                    <label>Foundation Term</label>
                    <input type="text" name="wlft" id="wlft" readonly ="true" value='<%Response.write rs("WorkingLiasionFoundationTerm") %>'/>
                     
   	                <br />  <br/><br/>  <br/><br/><br/>
					
                    <label>Working Liasion Concentration</label>
                    <input type="text" name="wlc" id="wlc" readonly ="true" value='<%Response.write rs("WorkingLiasionConcentration") %>'/>
   	                <label>Faculty Liasion Concentration</label>
                    <input type="text" name="flc" id="flc" readonly ="true" value='<%Response.write rs("FacultyLiasionConcentration") %>'/>                  
                    <label>Concentration Term</label>
                    <input type="text" name="wlct" id="wlct" readonly ="true" value='<%Response.write rs("WorkingLiasionConcentrationTerm") %>'/>
                    
   	                
                    
                   
                    <br/><br/><br/><br/><br/><br/>

                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/><br/><br/>
                        
                       <br/> <button type="submit" name="Button2" onclick="this.form.action='RemoveField.asp?UIN=<%Response.Write (UIN) %>'; this.forms.submit();" id="Button2" value='<% Response.write (UIN) %>'>Confirm to Remove Field</button>
         
                    <br/><br/>
                    </p>
				</form>
                </div>
               </div>
        
            </div>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</body>
</html>

