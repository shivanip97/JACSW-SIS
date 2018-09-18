<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UID = Request("Button1")
UIN = Request.QueryString("UIN")
set rs=Server.CreateObject("ADODB.recordset")
query="select * from Field1 where UID ='"& uid &"'"
rs.Open query,conn

SupervisorFullName = rs.Fields("FieldInstructor")
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<!--#include file="header.asp"-->
<title>SIS | Edit Agency</title>
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
        $('.gpa').mask('0.00');

        // Reset Checkbox values
        $('.checkboxField').each(function () {
            if ($(this).val() == "Y" || $(this).val() == "True") {
                $(this).attr('checked', true);
            }
            else {
                $(this).attr('checked', false);
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


    });
    function validate() {
        var shouldProceed = true;
        $('#agencyForm').find(':input:not(button)').each(function () {
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
    function getval(sel) {
        window.location = "https://socialwork.cc.uic.edu/SIS/EditAgency.asp?ID=" + Response.write(UIN);
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
            <!--<th align="center">Field Type</th>
            <th align="center">POE</th> -->
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
			<!--<td align="center"><div class="edit" id="<%= uin%> Field Type"> <% Response.write rs1("FieldType") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> POE"> <% Response.write rs1("POE") %></div></td> -->
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
</table></div>
            <div id="steps">
				<form id="agencyForm" method="post" action="AfterEditAgency.asp">

					<h3>Edit Agency Information</h3>
                    <br/> 
                   
                    <a href="ShowAgency.asp?UIN=<%Response.write(UIN) %>">Back to Show Placements</a> 
                    <br/> <br/>
                    <p>
                    <label>Agency Form</label>
                    <br/><br/><br/>
                        <%
                        set drs=Server.CreateObject("ADODB.recordset")
                        agency_query="select AgencyID from Agency1 where AgencyID= '"& rs.Fields("AgencyID") &"'"
                        drs.Open agency_query,conn1 
                          if not drs.EOF then
                          AgencyID = drs.Fields(0)
                                                         
                          end if
                          drs.close
                        %>
                    <label>AgencyID</label>
                   <input type="text" name="agencyID" id="agencyID" readonly="true" value ='<%Response.write(AgencyID)%>'/>
                        <br/><br/><br/> 
                    <label>Agency</label>
                    <textarea id="agency" name="agency" cols="70" rows="1"><%Response.write rs.Fields("Agency") %></textarea>
                   
                     <br/><br/><br/> 
                        
                    
                    <label>Term</label>
   	                <select name="term" id="term">
         			<option value="<%= rs.Fields("Term") %>"><%= rs.Fields("Term") %></option>
  					<option value="Spring 2013">Spring 2013</option>
					<option value="Summer 2013">Summer 2013</option>
                    <option value="Fall 2013">Fall 2013</option>
                    <option value="Spring 2014">Spring 2014</option>
                    <option value="Summer 2014">Summer 2014</option>
                    <option value="Fall 2014">Fall 2014</option>
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
                    <option value="<%= rs.Fields("FieldTypeYear") %>"><%= rs.Fields("FieldTypeYear") %></option>
                    <option value="Foundation">Foundation</option>
                    <option value="Concentration">Concentration</option>
                    <option value="Generalist">Generalist</option>
                    <option value="Specialization">Specialization</option>
                    
				    </select>
                     <br/><br/><br/>  <br/>
                    <label>Field Instructor</label>
                    <select name="fieldInst" id="fieldInst">
                     <%
                        query = "select SupervisorID,SupervisorFullName from Supervisor1 where AgencyID like '" & AgencyID & "' "
                        set drs = conn1.execute(query)
                         if drs.eof then%>
                        <option value="No Instructor Found" >No Instructor Found</option>
                         <% end if
                        do while not drs.eof %>
                        <option value="<%= drs.Fields(1) %>" ><%= drs.Fields(1) %></option>
                        <% drs.MoveNext
                        Loop
                    %>
                        <option value="<%= rs.Fields(11) %>" selected><%= rs.Fields("FieldInstructor") %></option>
                     </select>
                        <%
                        SupPquery = "select EmailAddress,Phone from Supervisor1 where SupervisorFullName = '"&rs.Fields("FieldInstructor")&"' "
                        set drs2 = conn1.execute(SupPquery)
                         if not drs2.eof then%>
                        <label>Phone</label>
                   <input type="text" name="Phone" class = "workphone" id="Phone" readonly="true" value ='<%=drs2.Fields(1)%>'/>
                        <label>Email</label>
                   <input type="text" name="Email" id="Email" readonly="true" value ='<%=drs2.Fields(0)%>'/>
                        <% end if
                    %>
                        <br/><br/><br/><br/>
                        
                   
                         
                    
                     <%
                        set drs1=Server.CreateObject("ADODB.recordset")
                        agencyAdd_query="select AddressL1, AddressL2, City, State, Zip from AgencyAddress1 where AgencyID like '"& AgencyID &"'"
                        drs1.Open agencyAdd_query,conn1 
                          if not drs1.EOF then
                         AddressL1 = drs1.Fields(0)
                         AddressL2 = drs1.Fields(1)
                         City = drs1.Fields(2)
                         State = drs1.Fields(3)
                         Zip = drs1.Fields(4)                          
                          end if
                          drs1.close
                        %>   
                   <label>AddressL1</label>
                   <input type="text" name="addressL1" id="addressL1" readonly="true" value ='<%Response.write(AddressL1)%>'/>
                        <label>AddressL2</label>
                   <input type="text" name="addressL2" id="addressL2" readonly="true" value ='<%Response.write(AddressL2)%>'/>
                        
                        <label>City</label>
                   <input type="text" name="City" id="City" readonly="true" value ='<%Response.write(City)%>'/>
<br/><br/><br/>
                        <label>State</label>
                   <input type="text" name="State" id="State" readonly="true" value ='<%Response.write(State)%>'/>
                        
                        <label>Zip</label>
                   <input type="text" name="Zip" id="Zip" readonly="true" value ='<%Response.write(Zip)%>'/>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/><br/><br/>
                    <button type="submit" name="Submit" onclick="this.form.action='AfterEditAgency.asp?UIN=' + this.value; this.forms.submit();" value='<% Response.Write (UIN) %>'>Save Placement</button><br />
					<button type="submit" name="Submit1" onclick="this.form.action='RemoveAgencyConfirm.asp?UIN=<%Response.Write (UIN) %>'; this.forms.submit();" id="Button1" value='<% Response.write (UID) %>' ">Remove Placement</button><br /><br />
					<br/><br/>
                    </p>
				</form>
               </div>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</bod>
</html>
