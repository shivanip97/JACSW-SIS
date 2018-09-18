<!--#include file="Login_Check.asp"-->
<!--#include file="DBConn.asp"-->
<% 
ErrMsg = Request("ErrMsg")
AgencyID = Request("Button1")

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
  <!--#include file="header.asp"-->
<title>SIS | Add Agency</title>
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
		$('.zip').mask('00000');

		// Reset Checkbox values
		$('.checkboxField').each(function () {
			if ($(this).val() == "Y" || $(this).val() == "True" || $(this).val() == "1") {
				$(this).attr('checked', true);
			}
			else {
				$(this).attr('checked', false);
			}


		});

		$('.checkboxField').on('click', function () {
			if ($(this).is(":checked")) {
				$(this).attr('value', '1');
			} else {
				$(this).attr('value', '0');
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
	<h3>Add MSW Agency Information</h3>
					 <br/>
					<a href="ShowAllAgency.asp">Back to Show Agency</a> 
					<br/> <br/>
					<% If Session("Username") = "test_a"  or Session("Username") = "tmorri3" or Session("Username") = "apradh6" or Session("Username") = "carrasc1" or Session("Username") = "bc1972" or Session("Username") = "nrosal1" or Session("Username") = "ktboyd" or Session("Username") = "chanze1" or Session("Username") = "fisherj" or Session("Username") = "melka1" or Session("Username") = "tashamc" or Session("Username") = "ajohns5" Then %>
					<button style="border:none;outline:none;-moz-border-radius: 10px;-webkit-border-radius: 10px;font-weight:bold;margin: 0px auto;clear:both;padding: 7px 25px;font-size:22px;display: block; background:#4797ED;font-family:Century Gothic, Helvetica, sans-serif;" type="submit" name="Button1" onclick="agencyForm.action='AfterAddNewAgency.asp' ; agencyForm.submit();" id="Button1" value=''>Save Agency</button><br /><br />
					<% End If %>
					
	<div id="wrapper">
	 <div id="navigation" style="display: none;">
					<ul>
						<li><a href="#">Agency Details</a></li>
					</ul>
			  </div>
	<div id="steps" align="center">
				<form id="agencyForm" method="post" action="AfterAddNewAgency.asp">
				
				<fieldset class="step">
				<legend></legend>

				<p>
					<br/><br/><br/>
					<label>Agency</label>
					<textarea id="Agency"  name="Agency" cols="70" rows="1"></textarea>
									   
					<br/><br/><br/><br/>
					<label>In Use</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="InUseFoundation" id="InUseFoundation" class="checkboxField" />
					<label class="clearWidth">Generalist (formerly Foundation)</label>
					<input type="checkbox"  style="margin:0;width:20px;height:20px;" name="InUseMH" id="InUseMH" class="checkboxField"  />
					<label class="clearWidth">MH</label>
					<input type="checkbox"  style="margin:0;width:20px;height:20px;" name="InUseCF" id="InUseCF" class="checkboxField"  />
					<label class="clearWidth">CF</label>
					<input type="checkbox"  style="margin:0;width:20px;height:20px;" name="InUseCHUD" id="InUseCHUD" class="checkboxField"  />
					<label class="clearWidth">OCP</label>
					<input type="checkbox" style="margin:0;width:20px;height:20px;" name="InUseSCH" id="InUseSCH" class="checkboxField" />
					<label class="clearWidth">SCH</label>
					
					
					
					<br/><br/><br/>
					<label>Note</label>
					
					<textarea id="Note" name="Note" cols="70" rows="5"></textarea>
					<br/><br/><br/><br/>
					
				 
	 
					<label>Agency Address Line 1</label>
					<input type="text" name="AddressL1"  id="AddressL1" />
					<label>Agency Address Line 2</label>
					<input type="text" name="AddressL2"  id="AddressL2" />  
					<label>Agency Phone</label>
					<input type="text" name="AgencyPhone" class="workphone" id="AgencyPhone"  />                                                         
					
					<br/><br/><br/><br /> 
					<label>City</label>
					<input type="text" name="City" id="City" />                
					<label>State</label>
					<input type="text" name="State" id="State" />
					<label>ZipCode</label>
					<input type="text" name="Zip" class="zip"  id="Zip" />
					
					   <br/><br/><br/>
					 <br/><br/><br/>
					<label>Agency Contact</label>
					<input type="text" name="Person" id="Person" />
					<label>Agency Contact Phone</label>
					<input type="text" name="AgencyContactPhone" class="workphone" id="AgencyContactPhone"  />
					<label>Agency Contact Email</label>
					<input type="text" name="Email" id="Email"  />
					<br/><br/><br/><br/>
					   

					
				   <label>School District</label>
					<input type="text" name="SchoolDistrict"  id="SchoolDistrict" /> 
				   <!-- <label>Agency Contact Phone Ext</label>
					<input type="text" name="Ext"  id="Ext"  />-->
					<br/><br/><br/><br/>
					<label>Website Address</label>
					<textarea id="WebsiteAddress"  name="WebsiteAddress" cols="70" rows="1"></textarea>
					
					 <br/><br/><br/><br/>
					<label>Description</label>
					<textarea id="Description"  name="Description" cols="70" rows="5"></textarea>  
					
					</p>
					</fieldset>
			

					
				   

				</form>
				</div>
			   </div>
			   
			</div>
			<br/>
			<!--#include file="footer.asp"-->
</body>
</html>
