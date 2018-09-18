<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
UID = Request("Submit1")
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
            <th align="center">Field Type</th>
            <th align="center">POE</th>
            <th align="center">Faculty Liasion Foundation</th>
            <th align="center">Faculty Liasion Concentration</th>
            <th align="center">Working Liasion Foundation</th>
            <th align="center">Working Liasion Concentration</th>
            <th align="center">Foundation Term</th>
            <th align="center">Concentration Term</th>       
		</tr>
	</thead>
    <tbody>
	 
        <tr>
			<td align="center"><% Response.write rs1("FieldType") %></td>
            <td align="center"><div class="edit" id="<%= uin%> POE"> <% Response.write rs1("POE") %></div></td>
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
                        agency_query="select AgencyID from Agency1 where Agency like '"& rs.Fields("AgencyID") &"'"
                        drs.Open agency_query,conn1 
                          if not drs.EOF then
                          AgencyID = drs.Fields(0)
                                                         
                          end if
                          drs.close
                        %>
                    <label>Agency ID</label>
                   <input type="text" name="agencyID" id="agencyID" readonly="true" value ='<%Response.write(AgencyID)%>'/>
                        <br/><br/><br/> 
                    <label>Agency</label>
                    <textarea id="agency" name="agency" readonly="true" cols="70" rows="1"><%Response.write Replace(rs.Fields("Agency"),"'","''") %></textarea>
                   
                     <br/><br/><br/> 
                        
                     
                        <label>Term</label>
                        <input type="text" name="term" id="term" readonly="true" value ='<%= rs.Fields(10)%>'/>
   	                
                        <label>Field Type Year</label>
                        <input type="text" name="fty" id="fty" readonly="true" value ='<%= rs.Fields(12)%>'/>
   	                
                        <br/><br/><br/>
                    <label>Field Instructor</label>
                    
                    
                        <input type="text" name="fieldInst" id="fieldInst" readonly="true" value ='<%= rs.Fields(11) %>'/>
                       
                       
                        <%
                        SupPquery = "select EmailAddress,Phone from Supervisor1 where SupervisorFullName = '"&rs.Fields(11)&"' "
                        set drs2 = conn1.execute(SupPquery)
                         if not drs2.eof then%>
                        <label>Phone</label>
                   <input type="text" name="Phone" class = "homephone" id="Phone" readonly="true" value ='<%=drs2.Fields(1)%>'/>
                        <label>Email</label>
                   <input type="text" name="Email" id="Email" readonly="true" value ='<%=drs2.Fields(0)%>'/>
                        <% end if
                    %>
                        <br/><br/><br/>
                        
                   
                         
                    
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
                    
					<button type="submit" name="Submit1" onclick="this.form.action='RemoveAgency.asp?UIN=<%Response.Write (UIN) %>'; this.forms.submit();" id="Submit1" value='<% Response.Write (UID)  %>'>Confirm to Remove Placement</button><br /><br />
					<br/><br/>
                    </p>
				</form>
               </div>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</bod>
</html>
