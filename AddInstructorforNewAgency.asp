﻿<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
AgencyID = Request("AgencyID")
    
set rs=Server.CreateObject("ADODB.recordset")
query="select * from Agency1 a full join AgencyAddress1 b on a.AgencyID = b.AgencyID full join AgencyNotes1 c on b.AgencyID = c.AgencyId  where a.AgencyID ='"& AgencyID &"'"
rs.Open query,conn1
if (IsNull(rs("ContactSCH"))  or IsEmpty( rs("ContactSCH")) ) then ContactSCH = "" else ContactSCH = rs("ContactSCH") End If
if (IsNull(rs("ContactMH")) or   IsEmpty(rs("ContactMH"))) then ContactMH = "" else ContactMH = rs("ContactMH")  End If   
if (IsNull(rs("ContactFoundation")) or IsEmpty(rs("ContactFoundation"))) then ContactFoundation = "" else ContactFoundation = rs("ContactFoundation") End If    
if (IsNull(rs("ContactCF"))or   IsEmpty(rs("ContactCF"))) then ContactCF = "" else ContactCF = rs("ContactCF") End If    
if (IsNull(rs("ContactCHUD")) or  IsEmpty(rs("ContactCHUD"))) then ContactCHUD = "" else ContactCHUD = rs("ContactCHUD") End If    
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | Edit Agency</title>
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

        $('.checkboxField').on('click', function () {
            if ($(this).is(":checked")) {
                $(this).attr('value', 'Y');
            } else {
                $(this).attr('value', 'N');
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

        $('#withdrawn').change(function () {
            $("#withdraw_reason").prop("disabled", !$(this).is(':checked'));
        });

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
    <h3>Edit MSW Agency Information</h3>
                     <br/>
                    <a href="ShowAllAgency.asp">Back to Show Agency</a> 
                    <br/> <br/>
                   
                    
    <div id="wrapper">
     <div id="navigation" style="display: none;">
                    <ul>
                        <li><a href="#">Agency Details</a></li>
                        <li><a href="#">Field Instructor</a></li>
                        
                    </ul>
              </div>
    <div id="steps" >
				<form id="agencyForm" method="post" action="SaveFacultyAfterAddAddressSISAgency.asp">
                
                <fieldset class="step">
                <legend></legend>

                <p>
                    <br/><br/><br/>
                    <label>Agency ID</label>
					<input type="text" name="AgencyID" required id="AgencyID" value='<%Response.write(AgencyID) %>' readonly=true/>   
                                  
                    <br/><br/><br/>

                    <label>Agency</label>
                    <textarea id="Agency" disabled="disabled" name="Agency" cols="70" rows="1"><%Response.write rs("Agency") %></textarea>
					<br/><br/><br/>
                    <label>In Use</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="InUseFoundation" id="InUseFoundation" class="checkboxField" value='<%Response.write rs("InUseFoundation") %>' />
                    <label class="clearWidth">Foundation</label>
					<input type="checkbox"  disabled="disabled" style="margin:0;width:20px;height:20px;" name="InUseMH" id="InUseMH" class="checkboxField" value='<%Response.write rs("InUseMH") %>' />
                    <label class="clearWidth">MH</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="InUseCF" id="InUseCF" class="checkboxField" value='<%Response.write rs("InUseCF") %>' />
                    <label class="clearWidth">CF</label>
					<input type="checkbox" disabled="disabled" style="margin:0;width:20px;height:20px;" name="InUseCHUD" id="InUseCHUD" class="checkboxField" value='<%Response.write rs("InUseCHUD") %>' />
                    <label class="clearWidth">CHUD</label>
					<input type="checkbox"  disabled="disabled" style="margin:0;width:20px;height:20px;" name="InUseSCH" id="InUseSCH" class="checkboxField" value='<%Response.write rs("InUseSCH") %>' />
                    <label class="clearWidth">SCH</label>

                    
                    <br/><br/><br/>
                     <br/><br/><br/>
                    <label>Agency Address Line 1</label>
					<input type="text" name="AddressL1"  id="AddressL1" value='<%Response.write rs("AddressL1") %>' readonly=true/>
                    <label>Agency Address Line 2</label>
					<input type="text" name="AddressL2"  id="AddressL2" value='<%Response.write rs("AddressL2") %>' readonly=true/>                                                           
                    <label>City</label>
                    <input type="text" name="City" id="City" value='<%Response.write trim(rs("City")) %>' readonly=true/>
                    <br/><br/><br/><br /> 
                                      
                    <label>State</label>
					<input type="text" name="State" id="State" value='<%Response.write trim(rs("State")) %>' readonly=true/>
                    <label>ZipCode</label>
					<input type="text" name="Zip" class="zip"  id="Zip" value='<%Response.write trim(rs("Zip")) %>' readonly="true" />
                    <label>Agency Phone</label>
					<input type="text" name="AgencyPhone"  id="AgencyPhone" value='<%Response.write rs("Phone") %>' readonly="true"/>
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
                    <input type="text" name="Person" id="Person" value='<%Response.write trim(ContactCF) %>' readonly="true"/>
                    <%elseif (ContactCHUD <> "") then%>
                    <input type="text" name="Person" id="Person" value='<%Response.write trim(ContactCHUD) %>' readonly="true"/>
                    <%End If %>
                    <label>Agency Contact Phone</label>
					<input type="text" name="AgencyContactPhone" class="homephone" id="AgencyContactPhone" value='<%Response.write trim(rs("AgencyContactPhone")) %>' readonly="true"/>
                    <label>Agency Contact Email</label>
					<input type="text" name="Email" id="Email" value='<%Response.write trim(rs("Email")) %>' readonly="true"/>
                   
                    
                   
                    <br/><br/><br/><br/>
                    <label>School District</label>
					<input type="text" name="SchoolDistrict"  id="SchoolDistrict" value='<%Response.write rs("SchoolDistrict") %>' readonly="true"/> 
                     <br/><br/><br/><br/>
                    <label>Website Address</label>
                    <textarea id="WebsiteAddress" disabled="disabled" name="WebsiteAddress" cols="70" rows="1"><%Response.write rs("WebsiteAddress") %></textarea>
                      <br/><br/><br/><br/>
                    <label>Description</label>
                    <textarea id="Description"  disabled="disabled" name="Description" cols="70" rows="5"><%Response.write rs("Description") %></textarea> 
                    </p><br/><br/>
                    
                     </fieldset>
                    <%rs.close %>

                      <div id="fieldinstructor" >
                    <fieldset class="step">
                    <legend></legend>
                
                                          
                    <p>
                    <br/>
                    <label>Field Instructor</label>
                    <input type="text" name="SupervisorFullName" id="SupervisorFullName" value= '' />
                    <label>Email</label>
					<input type="text" name="EmailAddress" id="EmailAddress" value=''/>
                    <br/><br/><br/><br /> 
                                  
                    
                    <label>Work Phone</label>
					<input type="text" name="SPhone"  class="workphone" id="SPhone" value='' />
                    <label>Cell Phone</label>
					<input type="text" name="CellPhone" class="homephone"  id="CellPhone" value='' /> 
                   
                    <br/><br/><br/>
                    <br/><br/><br/><br/>
                    </p>

                    <br/><br/>
                     <button type="submit" name="Submit" onclick="this.form.action='SaveFacultyAfterAddAddressSISAgency.asp?AgencyID=' + this.value; this.forms.submit();" value='<%Response.write(AgencyID) %>'>Save</button>
                    <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    
                      </fieldset>
                   
				</form>
                </div>
               </div>
               
			</div>
            
            <!--#include file="footer.asp"-->
</body>
</html>