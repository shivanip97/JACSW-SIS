<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
AgencyID = Request("Button1")
AgencyID=Request.QueryString("AgencyID")
	
set rs=Server.CreateObject("ADODB.recordset")
query="select * from Agency1 a full join AgencyAddress1 b on a.AgencyID = b.AgencyID full join AgencyNotes1 c on b.AgencyID = c.AgencyId  where a.AgencyID ='"& AgencyID &"'"
rs.Open query,conn1
if (IsNull(rs("ContactSCH"))  or IsEmpty( rs("ContactSCH")) ) then ContactSCH = "" else ContactSCH = rs("ContactSCH") End If
if (IsNull(rs("ContactMH")) or   IsEmpty(rs("ContactMH"))) then ContactMH = "" else ContactMH = rs("ContactMH")  End If   
if (IsNull(rs("ContactFoundation")) or IsEmpty(rs("ContactFoundation"))) then ContactFoundation = "" else ContactFoundation = rs("ContactFoundation") End If    
if (IsNull(rs("ContactCF"))or   IsEmpty(rs("ContactCF"))) then ContactCF = "" else ContactCF = rs("ContactCF") End If    
if (IsNull(rs("ContactCHUD")) or  IsEmpty(rs("ContactCHUD"))) then ContactCHUD = "" else ContactCHUD = rs("ContactCHUD") End If  
if (IsNull(rs("SchoolDistrict")) or  IsEmpty(rs("SchoolDistrict"))) then ContactCHUD = "" else ContactCHUD = rs("SchoolDistrict") End If  
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
 

<title>SIS | View Agency</title>
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
	<h3>View MSW Agency Information</h3>
					 <br/>
					<a href="ShowAllAgency.asp">Back to Show Agency</a> 
				 <br />
		 <br />
		 <br />
	<div id="wrapper">
	 <div id="navigation" style="display: none;">
					<ul>
						<li><a href="#">Agency Details</a></li>
						<li><a href="#">Field Instructor</a></li>
			  
					</ul>
			  </div>
	<div id="steps" align="center">
				<form id="agencyForm" method="post" action="EditSISAgency.asp">
				
				<fieldset class="step">
				<legend></legend>

				<p>
					<br/><br/><br/>
					
					<label>Agency</label>
					<textarea id="Agency" disabled="disabled" name="Agency" cols="70" rows="1"><%Response.write rs("Agency") %></textarea>
					
					<br/><br/><br/>
					
					<label>In Use</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="InUseFoundation" id="InUseFoundation" class="checkboxField" value='<%Response.write rs("InUseFoundation") %>' />
					<label class="clearWidth">Generalist (formerly Foundation)</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="InUseMH" id="InUseMH" class="checkboxField" value='<%Response.write rs("InUseMH") %>' />
					<label class="clearWidth">MH</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="InUseCF" id="InUseCF" class="checkboxField" value='<%Response.write rs("InUseCF") %>' />
					<label class="clearWidth">CF</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="InUseCHUD" id="InUseCHUD" class="checkboxField" value='<%Response.write rs("InUseCHUD") %>' />
					<label class="clearWidth">OCP</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="InUseSCH" id="InUseSCH" class="checkboxField" value='<%Response.write rs("InUseSCH") %>' />
					<label class="clearWidth">SCH</label>    
					<br/><br/><br/>
					 
					<label>Note</label>
					
					<textarea id="Note" disabled="disabled" name="Note" cols="70" rows="5"><%Response.write rs("Note") %></textarea>
					<br/><br/><br/><br/>
	 
					<label>Agency Address Line 1</label>
					<input type="text" name="AddressL1"  id="AddressL1" value='<%Response.write rs("AddressL1") %>' readonly=true/>
					<label>Agency Address Line 2</label>
					<input type="text" name="AddressL2"  id="AddressL2" value='<%Response.write rs("AddressL2") %>' readonly=true/>  
					 <label>Agency Phone</label>
					<input type="text" name="City" id="Text1" value='<%Response.write trim(rs("Phone")) %>' readonly=true/>                                                         
					
					<br/><br/><br/><br /> 
					<label>City</label>
					<input type="text" name="City" id="City" value='<%Response.write trim(rs("City")) %>' readonly=true/>                  
					<label>State</label>
					<input type="text" name="State" id="State" value='<%Response.write trim(rs("State")) %>' readonly=true/>
					<label>ZipCode</label>
					<input type="text" name="Zip" class="zip"  id="Zip" value='<%Response.write trim(rs("Zip")) %>' readonly=true/>
				   
					   <br/><br/><br/>
					 <br/><br/><br/>

					<label>Agency Contact</label>
					
					<%if (ContactFoundation <> "") then%>
					<input type="text" name="Person" id="Person" value='<%Response.write trim(ContactFoundation) %>' readonly="true"/>
					<%elseif (ContactSCH <> "") then %>
					<input type="text" name="Person" id="Person" value='<%Response.write trim(ContactSCH) %>' readonly="true"/>
					<%elseif (ContactMH <> "") then %>
					<input type="text" name="Person" id="Person" value='<%Response.write trim(ContactMH) %>' readonly="true"/>
					<%elseif (ContactCF <> "") then %>
					<input type="text" name="Person" id="Person" value='<%Response.write trim(ContactCF) %>' readonly="true" />
					<%elseif (ContactCHUD <> "") then%>
					<input type="text" name="Person" id="Person" value='<%Response.write trim(ContactCHUD) %>' readonly="true"/>
					<%else%>
					<input type="text" name="Person" id="Person" value='' readonly="true"/>
					<%End If %>
					
					
					<label>Agency Contact Phone</label>
					<input type="text" name="Phone" class="workphone" id="Phone" value='<%Response.write trim(rs("AgencyContactPhone")) %>' readonly="true"/>
					<label>Agency Contact Email</label>
					<input type="text" name="Email" id="Email" value='<%Response.write trim(rs("Email")) %>' readonly="true"/>
					 <br/><br/><br/><br/>
					<label>School District</label>
					<input type="text" name="SchoolDistrict"  id="SchoolDistrict" value='<%Response.write trim(rs("SchoolDistrict")) %>' readonly=true/> 

					<br/><br/><br/><br/>
					<label>Website Address</label>
					<textarea id="WebsiteAddress" disabled="disabled" name="WebsiteAddress" cols="50" rows="1"><%Response.write rs("WebsiteAddress") %></textarea> 
					
					<br/><br/><br/><br/>
					<label>Description</label>
					<textarea id="Description" disabled="disabled" name="Description" cols="70" rows="5"><%Response.write rs("Description") %></textarea>  
					
					</p>
					</fieldset>
					<%rs.close %>

					  <div id="fieldinstructor" align="center">
					<fieldset class="step">
					<legend></legend>
				
					<% 
ErrMsg = Request("ErrMsg")
AgencyID = Request("Button1")
AgencyID=Request.QueryString("AgencyID")
set rs=Server.CreateObject("ADODB.recordset")
Supquery="select a.SupervisorID as supID , * from Supervisor1 a full join SupervisorNotes1 b on a.SupervisorID = b.SupervisorID where a.AgencyID ='"& AgencyID &"'"
rs.Open Supquery,conn1
						 If rs.EOF Then
					  
					Else
					'if there are records then loop through the fields
					i=1
					Do While NOT rs.Eof 
%>
					   
					<p>
					<br/>
					<label>Field Instructor</label>
					<input type="text" name="SupervisorFullName" readonly="true" id="SupervisorFullName" value= "<% Response.write trim(rs("SupervisorFullName")) %>" />
					<label>Email</label>
					<input type="text" name="EmailAddress" readonly="true" id="EmailAddress" value="<%Response.write trim(rs("EmailAddress")) %>" />
					<br/><br/><br/><br /> 
				  
					<label>Work Phone</label>
					<input type="text" name="FPhone"  class="workphone" readonly="true" id="FPhone" value='<%Response.write trim(rs("Phone")) %>' />
					<label>Cell Phone</label>
					<input type="text" name="CellPhone" class="homephone" readonly="true" id="CellPhone" value='<%Response.write trim(rs("CellPhone")) %>' /> 
				

					
						<br/><br/><br/>
					
					
			   
					</p>
				  
				   <%
					i=i+1
					rs.MoveNext    
					Loop
					End If

 rs.close

%>
				   
				  
			
					

					<strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
					<br/>
					 
					</fieldset>
				   

				</form>
				</div>
			   </div>
			   
			</div>
			<br/>
	   <br/> 
					<% If Session("Username") = "test_a" or Session("Username") = "tmorri3" or Session("Username") = "apradh6" or Session("Username") = "carrasc1" or Session("Username") = "bc1972" or Session("Username") = "nrosal1" or Session("Username") = "ktboyd" or Session("Username") = "chanze1" or Session("Username") = "fisherj" or Session("Username") = "melka1" or Session("Username") = "tashamc" or Session("Username") = "ajohns5" Then %>
					<button style="border:none;outline:none;-moz-border-radius: 10px;-webkit-border-radius: 10px;font-weight:bold;margin: 0px auto;clear:both;padding: 7px 25px;font-size:22px;display: block; background:#4797ED;font-family:Century Gothic, Helvetica, sans-serif;" type="submit" name="Button1" onclick="agencyForm.action='EditSISAgency.asp?AgencyID=' + this.value; agencyForm.submit();" value='<% Response.write(AgencyID)%>'>Edit Agency</button><br /><br />
					<% End If %>
					
			<!--#include file="footer.asp"-->
</body>
</html>
