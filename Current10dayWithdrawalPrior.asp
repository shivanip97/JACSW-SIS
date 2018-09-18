<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Login_Check.asp"-->
<!--#include file="fpdf.asp"-->
<!--#include file="DBconn.asp"-->

<link rel="stylesheet" href="css/tabstyle.css" type="text/css" />
    <script type="text/javascript" src="jquery/jquery-1.9.0.js"></script>
    <script type="text/javascript" src="jquery/filtertable.js"></script>
    <script type="text/javascript" src="jquery/jquery.jeditable.mini.js"></script>
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

	    function getval(sel) {
	        window.location = "https://socialwork.cc.uic.edu/SIS/Current10dayWithdrawalReport.asp?ID=" + sel.value;
	    }
 	</script>
<h4><a href="ShowCurrentStudents.asp">Home</a> |<a href="MSWCurrentReports.asp">Current Student Reports</a>  | <a href="logout.asp">Log Out</a></h4>
<div align="center"> 
 <label style="font-size: 1.17em;font-weight: bold;">Select Start Date</label>
    
      <select name="date" id="date" onchange="getval(this);">
          <option value="2016-08-08">2016-08-08</option>
          <option value="2016-08-09">2016-08-09</option>
          <option value="2016-08-10">2016-08-10</option>
          <option value="2016-08-11">2016-08-11</option>
          <option value="2016-08-12">2016-08-12</option>
          <option value="2016-08-13">2016-08-13</option>
          <option value="2016-08-14">2016-08-14</option>
          <option value="2016-08-15">2016-08-15</option>
          <option value="2016-08-16">2016-08-16</option>
          <option value="2016-08-17">2016-08-17</option>
          <option value="2016-08-18">2016-08-18</option>
          <option value="2016-08-19">2016-08-19</option>
          <option value="2016-08-20">2016-08-20</option>
          <option value="2016-08-21">2016-08-21</option>
          <option value="2016-08-22">2016-08-22</option>
          <option value="2016-08-23">2016-08-23</option>
          <option value="2016-08-24">2016-08-24</option>
          <option value="2016-08-25">2016-08-25</option>
          <option value="2016-08-26">2016-08-26</option>
          <option value="2016-08-27">2016-08-27</option>
          <option value="2016-08-28">2016-08-28</option>
          <option value="2016-08-29">2016-08-29</option>
          <option value="2016-08-30">2016-08-30</option>
          <option value="2016-08-31">2016-08-31</option>          
                    </select> 
      <br /><br />
    </div>

