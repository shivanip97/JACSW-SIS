<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UIN = Request("Button")
    
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | Add New Field</title>
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
            alert('Please Complete form by filling in fields highlighted in Red.')
        }

        return shouldProceed;
    }

    function getval(sel) {
        window.location = "https://socialwork.cc.uic.edu/SIS/AddAgency.asp?ID="+Response.write(UIN);
    }
 	</script>
 <style type="text/css">
		table {
			text-align: left;
			font-size: 12px;
			font-family: verdana;
			background: #c0c0c0;
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
    <div id="content" align=center>
        <br/><br/>
                    <p><label for="UIN">UIN: </label><strong><font color="#000000"><% Response.write(UIN) %></font></strong></p>
        
        <%
                    set rs2=Server.CreateObject("ADODB.recordset")
					course_query1="select FirstName, LastName from CurrentStudents where UIN = '"& UIN & "'"
					rs2.Open course_query1,conn
                    If rs2.EOF Then
                   Else
            %>
                    <p><label for="Name:">Name: </label><strong><font color="#000000"><% Response.Write rs2("FirstName") %> &nbsp <% Response.Write rs2("LastName") %> </font></strong></p>
                    
                  <% End If %>
        
 <br/>
         <%
                    set rs1=Server.CreateObject("ADODB.recordset")
					course_query="select * from Field1 where UIN like '"& UIN & "'"
					rs1.Open course_query,conn
                    %>
                    
          <div>
    <table id="studentsTable">
	<thead>

		<tr>
          <!--  <th align="center">Field Type</th>
            <th align="center">POE</th>-->
            <th align="center">Faculty Liasion Generalist</th>
            <th align="center">Faculty Liasion Specialization</th>
            <th align="center">Working Liasion Generalist</th>
            <th align="center">Working Liasion Specialization</th>
            <th align="center">Generalist Term</th>
            <th align="center">Specialization Term</th>
		</tr>
	</thead>
    <tbody>
	 
        <tr>
			<!--<td align="center"><% Response.write rs1("FieldType") %></td>
            <td align="center"><div class="edit" id="<%= uin%> POE"> <% Response.write rs1("POE") %></div></td>-->
            <td align="center"><div class="edit" id="<%= uin%> FacultyLiasionFoundation"> <% Response.write rs1("FacultyLiasionFoundation") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> FacultyLiasionConcentration"> <% Response.write rs1("FacultyLiasionConcentration") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> WorkingLiasion Foundation"> <% Response.write rs1("WorkingLiasionFoundation") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> WorkingLiasion Concentration"> <% Response.write rs1("WorkingLiasionConcentration") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> Foundation Term"> <% Response.write rs1("WorkingLiasionFoundationTerm") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> Concentration Term"> <% Response.write rs1("WorkingLiasionConcentrationTerm") %></div></td>

        </tr>
		  
      
	</tbody>
    <% rs1.Close
                Set rs1=Nothing
                
                %>
</table></div></br>
         <table id="agencyTable">
	<thead>
    
		<tr>
            <th align="center">Agency</th>
            
            <th align="center">Term</th>
            <th align="center">Field Type Year</th>
            <th align="center">Field Instructor</th>
            
		</tr>
	</thead>
    <tbody>
     <%
					set rs=Server.CreateObject("ADODB.recordset")
					course_query="select f.UID as UID,  f.Term as Term,f.FieldTypeYear as FieldTypeYear,f.Agency as Agency, f.FieldInstructor as FieldInstructor FROM Field1 f where f.UIN = '"& UIN & "' and Agency IS NOT NULL order by Agency"
					rs.Open course_query,conn 
                    If rs.EOF Then
                    Response.write("No Agencies Found.")
                    Else
                    Do While NOT rs.Eof  
                    uid = rs("UID")
           %>
		<tr>
            <td align="center"><div id="<%= uin%> Agency"> <% Response.write rs("Agency") %></div></td>
			
            <td align="center"><div class="edit" id="<%= uin%> Term"> <% Response.write rs("Term") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> FieldTypeYear"> <% Response.write rs("FieldTypeYear") %></div></td>
            <td align="center"><div id="<%= uin%> FieldInstructor"> <% Response.write rs("FieldInstructor") %></div></td>
            
         </tr>
		 <%   rs.MoveNext 
             Loop
             End If %>
	</tbody>
    <% rs.Close
                Set rs=Nothing
                conn.Close
                Set conn=Nothing
                %>
</table>
        
        <br/>
                    <a href="ShowAgency.asp?UIN=<%Response.write(UIN) %>">Back to Show Placements</a> 
                    <br/> <br/>
       
     
            <div id="steps">
				<form id="agencyForm" method="post" action="AfterAddAgency.asp">
					<h3>Add New Agency</h3>
                     
                    <p>
                    <label>Add Agency Form</label>
                    <br/><br/><br/>
                    <label>Agency</label>
                    <select name="agency" id="agency" style="width:600px" onchange="getval(this);">
                    <%
                        query = "select a.AgencyID,a.Agency,b.AddressL1 from Agency1 a, AgencyAddress1 b where a.AgencyID = b.AgencyID order by Agency"
                        
                        set drs = conn1.execute(query)
                        
                        %>
                        <option value=""></option>
                        <% 
                        do while not drs.eof
                        %>
                        <option value="<%Response.write drs("AgencyID") %>"><%= drs.Fields(1)  %>&nbsp&nbsp<%="("%><%= drs.Fields(2)  %><%=")"%></option>
                        <%     
                        drs.MoveNext
                        Loop
                    %>
                     </select>

                        

                     
                        <br/><br/><br/>
                    
                       
                    
                    <label>Term</label>
   	                <select name="term" id="term">
         			<option value="0">-- Select --</option>
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
                    <label>Field Type Year</label>
   	                <select name="fty" id="fty">
                    <option value="0">-- Select --</option>
                    <option value="Foundation">Foundation</option>
                    <option value="Concentration">Concentration</option>
                    <option value="Generalist">Generalist</option>
                    <option value="Specialization">Specialization</option>
                    
				    </select>
                    <br/><br/><br/>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/><br/><br/>
					<button type="submit" name="Submit" id="Submit" onclick="this.form.action='AfterAddAgency.asp?UIN=' + this.value; this.forms.submit();" value='<%Response.Write (UIN) %>'>Add Agency</button>
					<br/><br/>
                    </p>
				</form>
               </div>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</body>
</html>
