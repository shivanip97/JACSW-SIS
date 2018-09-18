<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
AgencyID=Session("AgencyID")
ErrMsg = Request("ErrMsg")
UIN = Request.QueryString("UIN")
set rs=Server.CreateObject("ADODB.recordset")
query="select Max(UID) as uid from Field1 where UIN= '"& UIN &"'"
rs.Open query,conn  
    uid = rs.Fields(0)
   set ars1=Server.CreateObject("ADODB.recordset")  
getagencyquery = "select Agency,AgencyID from Field1 where UID ='"& uid &"'"
    ars1.Open getagencyquery,conn 
    agency = Replace(ars1.Fields(0), "'", "''")
    agency1= Replace(ars1.Fields(0), "'", "''''")
    aid=AgencyID
set rs1=Server.CreateObject("ADODB.recordset")
agencyquery="select * from Field1 where UID ='"& uid &"'"
rs1.Open agencyquery,conn    
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
    var entityMap = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;',
        '/': '&#x2F;',
        '`': '&#x60;',
        '=': '&#x3D;'
    };

    function escapeHtml(string) {
        return String(string).replace(/[&<>"'`=\/]/g, function (s) {
            return entityMap[s];
        });
    }
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
        window.location = "https://socialwork.cc.uic.edu/SIS/AddInstructorAgency.asp?ID=" + Response.write(UIN);
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
                    set rs3=Server.CreateObject("ADODB.recordset")
					course_query="select * from Field1 where UIN like '"& UIN & "'"
					rs3.Open course_query,conn
                    %>
                    
          <div>
    <table id="studentsTable">
	<thead>

		<tr>
            <th align="center">Field Type Year</th>
            
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
			<td align="center"><% Response.write rs3("FieldTypeYear") %></td>
           
            <td align="center"><div class="edit" id="<%= uin%> FacultyLiasionFoundation"> <% Response.write rs3("FacultyLiasionFoundation") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> FacultyLiasionConcentration"> <% Response.write rs3("FacultyLiasionConcentration") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> WorkingLiasion Foundation"> <% Response.write rs3("WorkingLiasionFoundation") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> WorkingLiasion Concentration"> <% Response.write rs3("WorkingLiasionConcentration") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> Foundation Term"> <% Response.write rs3("WorkingLiasionFoundationTerm") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> Concentration Term"> <% Response.write rs3("WorkingLiasionConcentrationTerm") %></div></td>
            
            
        </tr>
		  
      
	</tbody>
    <% rs3.Close
                Set rs3=Nothing
                
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
					set rs4=Server.CreateObject("ADODB.recordset")
					course_query="select  f.Term as Term,f.FieldTypeYear as FieldTypeYear,f.Agency as Agency, f.FieldInstructor as FieldInstructor, b.UID as UID FROM Field1 f, AddAgency1 b where f.Agency = b.Agency and f.FieldTypeYear = b.FieldTypeYear  and f.Term = b.Term and f.FieldInstructor = b.FieldInstructor and f.UIN = '"& UIN & "' order by Agency"
					rs4.Open course_query,conn 
                    If rs4.EOF Then
                    Response.write("No Agencies Found.")
                    Else
                    Do While NOT rs4.Eof  
                  
           %>
		<tr>
            <td align="center"><div id="<%= uin%> Agency"> <% Response.write rs4("Agency") %></div></td>
			
            <td align="center"><div class="edit" id="<%= uin%> Term"> <% Response.write rs4("Term") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> FieldTypeYear"> <% Response.write rs4("FieldTypeYear") %></div></td>
            <td align="center"><div id="<%= uin%> FieldInstructor"> <% Response.write rs4("FieldInstructor") %></div></td>
            
         </tr>
		 <%   rs4.MoveNext 
             Loop
             End If %>
	</tbody>
    <% rs4.Close
                Set rs4=Nothing
                
                %>
</table>
        
        <br/>
                    <a href="ShowAgency.asp?UIN=<%Response.write(UIN) %>">Back to Show Placement</a> 
                    <br/> <br/>
       <div id="wrapper">
     <div id="navigation" style="display: none;">
                    <ul>
                        <li><a href="#">Agency</a></li>
                        <li><a href="#"> Field Instructor</a></li>
                        
                    </ul>
              </div>
    <div id="steps" align="center">
   
				<form id="agencyForm" method="post" action="AfterAddInstructorAgency.asp?UIN=' + this.value; " value='<% Response.write (UIN) %>'>
                    <fieldset class="step">
               
                    <p>
                    <label>Add Agency Form</label>
                    <br/><br/><br/>
                         <%
                        set drs=Server.CreateObject("ADODB.recordset")
                        agency_query="select AgencyID from Agency1 where AgencyID like '"& aid &"'"
                        drs.Open agency_query,conn1 
                          if not drs.EOF then
                          AgencyID = drs.Fields(0)      
                          end if
                          drs.close
                        %>
                    <label>AgencyID</label>
                   <input type="text" name="agencyID" id="agencyid" readonly="true" value ='<%Response.Write(AgencyID)%>'/>
                        <br/><br/><br/> 
                    <label>Agency</label>
                    <textarea  name="agency" id="agency" disabled="disabled" cols="70" rows="1" ><%Response.Write ars1.Fields(0) %></textarea> 
                        <input type="hidden" id="uid1" name="uid1" value='<%Response.Write rs.Fields(0)%>' />

                                       
                     <br/><br/><br/> 
                        
                     
                    <label>Term</label>
                    <input id="term" name="term" readonly="true" type="text" value='<%=rs1.Fields("Term")%>' />
   	                <label>Field Type Year</label>
                    <input type="text" name="fty"   id="fty"  readonly="true" value ='<%=rs1.Fields("FieldTypeYear")%>'/>
                        
                   
                    
   	               
                    <br/><br/><br/>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    <br/><br/><br/>
					
					<br/><br/>
                    </p>
                        </fieldset>
                    <div id="application" align="center">
                    <fieldset class="step">
                    
                    <p>
                        <label>Add Instructor</label>
                    <br/><br/><br/>
                        <label>Field Instructor</label>
                    <select name="fieldInst" id="fieldInst">
                     <%
                        query = "select SupervisorID,SupervisorFullName from Supervisor1 where AgencyID like '" & aid & "' "
                        set drs = conn1.execute(query)
                         if drs.eof then%>
                        <option value="No Instructor Found" >No Instructor Found</option>
                         <% end if
                        do while not drs.eof %>
                        <option value="<%= drs.Fields(1) %>" ><%= drs.Fields(1) %></option>
                        <% drs.MoveNext
                        Loop
                    %>
                        
                        
                     </select>
                        <br/><br/><br/>
                   
                    <br/><br/><br/>
                        <button type="submit" name="Submit" onclick="this.form.action='AfterAddInstructorAgency.asp?UIN=' + this.value; this.forms.submit();" value='<%Response.Write (UIN) %>'>Add Instructor</button>
                        <br/><br/>
                    </p>
                        </fieldset>
                        </div>
				</form>
               </div>
               <!--#include file="footer.asp"-->
			</div>
            <br/>
</body>
</html>
