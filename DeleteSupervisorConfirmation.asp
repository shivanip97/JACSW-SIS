<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | Delete Faculty </title>
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
            if ($(this).val() == "Y") {
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
    <h3>Add MSW Agency Faculty Information</h3>
                     <br/>
                    <a href="ShowAllAgency.asp">Back to Show Agency</a> 
                    <br/> <br/>
                   
                    
    <div id="wrapper">
    
    <div id="steps" >
				<form id="agencyfacultyForm" method="post" action="DeleteFacultyAfterSelectingSISAgency.asp?AgencyID='<%Response.write(AgencyID) %>'">
                  
                
                <fieldset class="step">
                <legend></legend>

                
                        

 <% 
ErrMsg = Request("ErrMsg")
SupID = Request("SupervisorID")
    
set rs=Server.CreateObject("ADODB.recordset")
Supquery="select *, a.SupervisorID as supID from Supervisor1 a full join SupervisorNotes1 b on a.SupervisorID = b.SupervisorID where a.SupervisorID ='"& SupID &"'"
rs.Open Supquery,conn1
                         If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
%>
                       
                    <p>
                    <br/>
                    <label>Field Instructor</label>
                    <input type="text" name="SupervisorFullName" id="SupervisorFullName" value= '<% Response.write trim(rs("SupervisorFullName")) %>' readonly ="true"/>
                    <label>Email</label>
					<input type="text" name="EmailAddress" id="EmailAddress" value='<%Response.write trim(rs("EmailAddress")) %>' readonly=true/>
                    <label>SupervisorID</label>
					<input type="text" name="SupervisorID" id="SupervisorID" value='<%Response.write rs("supID") %>' readonly=true/>
                    <br/><br/><br/><br /> 
                                <input type="hidden" name="AgencyId"  id="Hidden1" value='<%Response.write rs("AgencyID") %>' readonly="true"/>   
                    
                    <label>Work Phone</label>
					<input type="text" name="Phone"  class="homephone" id="Phone" value='<%Response.write trim(rs("Phone")) %>' readonly=true/>
                    <label>Cell Phone</label>
					<input type="text" name="CellPhone" class="homephone"  id="CellPhone" value='<%Response.write trim(rs("CellPhone")) %>' readonly=true /> 
                    
                    
                    <br/><br/><br/><br/>
                       
                    <button type="submit" name="SubmitDelete" onclick="this.form.action='DeleteFacultyAfterSelectingSISAgency.asp?SupervisorId=' + this.value; this.forms.submit();" value='<% Response.write rs("supID") %>'>Confirm to be Deleted</button><br />
                    </p>
                  
                   <%
                    
                    End If

 rs.close

%>
                    
                        
                    
   <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
                    
                      </fieldset>
                   
				</form>
                </div>
               </div>
               
			</div>
            
            <!--#include file="footer.asp"-->
</body>
</html>