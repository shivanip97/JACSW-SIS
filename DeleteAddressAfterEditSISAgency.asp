<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | Delete Address</title>
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
    <h3>Add MSW Agency Address Information</h3>
                     <br/>
                    <a href="ShowAllAgency.asp">Back to Show Agency</a> 
                    <br/> <br/>
                   
                    
    <div id="wrapper">
    
    <div id="steps" >
				<form id="agencyaddressForm" method="post" action="DeleteAddressAfterSelectingSISAgency.asp?AgencyID='<%Response.write(AgencyID) %>'">
                
                <fieldset class="step">
                <legend></legend>

                <p>
                        

<% 
ErrMsg = Request("ErrMsg")
AgencyID = Request("AgencyID")
    
set rs=Server.CreateObject("ADODB.recordset")
Addressquery="select * from AgencyAddress where AgencyID ='"& AgencyID &"'"
rs.Open Addressquery,conn1
                         If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
%>


                    <label>Agency Address Line 1</label>
					<input type="text" name="AddressL1"  id="AddressL1" value='<%Response.write rs("AddressL1") %>' readonly ="true"/>
                    <label>Agency Address Line 2</label>
					<input type="text" name="AddressL2"  id="AddressL2" value='<%Response.write rs("AddressL2") %>' readonly ="true" />                                                           
                    <label>City</label>
                    <input type="text" name="City" id="City" value='<%Response.write rs("City") %>'  readonly ="true"/>
                    <br/><br/><br/><br /> 
                                      
                    <label>State</label>
					<input type="text" name="State" id="State" value='<%Response.write rs("State") %>' readonly ="true"/>
                    <label>ZipCode</label>
					<input type="text" name="Zip" class="zip"  id="Zip" value='<%Response.write rs("Zip") %>' readonly ="true"/>
                    <label>Agency AddressId</label>
					<input type="text" name="AddressId"  id="AddressId" value='<%Response.write rs("AddressId") %>' readonly="true"/>
                    <br/><br/><br/>
                    <button type="submit" name="SubmitDelete" onclick="this.form.action='DeleteAddressAfterSelectingSISAgency.asp?AddressId=' + this.value; this.forms.submit();" value='<% Response.write rs("AddressId") %>'>Delete</button><br />
                        
                     <br/><br/><br/>
                     <%
                    i=i+1
                    rs.MoveNext    
                    Loop
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