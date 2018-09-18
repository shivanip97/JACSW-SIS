<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>SIS | Prospective Students</title>
    <!--#include file="DBconn.asp"-->
    <!--#include file="Login_Check.asp"-->
    <!--#include file="header.asp"-->
    <% 
ErrMsg = Request("ErrMsg")
set rs=Server.CreateObject("ADODB.recordset")
query="select * from Students where UIN="&Application("currentUIN")
%>
    <link rel="stylesheet" href="css/tabstyle.css" type="text/css" media="screen" />
	<script type="text/javascript" src="jquery/jquery-1.8.2.min.js"></script>
    <script type="text/javascript" src="jquery/jquery.chromatable.js"></script>
    <script type="text/javascript" src="jquery/jquery-ui-1.8.2.custom.min.js"></script>
    <script type="text/javascript" src="jquery/sliding.form.js"></script>
    <script>

        $(document).ready(function () {
            if (!sessionStorage.getItem("tag")) {
                sessionStorage.setItem("tag", "1");
            }
            $('#navigation li:nth-child(' + parseInt(sessionStorage.getItem("tag")) + ') a').click();
            sessionStorage.removeItem("tag");

            $("#save").click(function (event) {
                sessionStorage.setItem("tag", "2");
            });

            $("#backStudents").click(function (event) {
                sessionStorage.removeItem("tag");
                window.location = "ProspectiveStudents.asp";
            });
        });

       
        
    </script>

        <style type="text/css">
		table {
			text-align: left;
			font-size: 12px;
			font-family: verdana;
			background: #c0c0c0;
			table-layout:fixed;
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
			height:50px;
			overflow: hidden;
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
    <div id="content">
        <h1>SIS - Prospective Students</h1>
        <br /><br />
          <a style="font-size:12pt;" href="UserHome.asp"><< HOME</a> | <a style="font-size:12pt;" href="Students.asp">Jump to Next Section >> </a>
        <br /><br /><br />
        <br />
        <div id="wrapper">
        <div id="navigation" style="display: none; visibility: hidden;">
                <ul>
                    <li><a href="#">Students</a></li>
                    <li><a href="#">Applicants</a></li>
                </ul>
            </div>
            <div id="steps" align="center">
                <form id="formElem" name="formElem" action="" method="post">
                <fieldset class="step">
                <legend></legend>
                <p>
                     <br/>
                         <label for="firstname">First Name</label>
                        <input name="firstname" type="text" class="required" id="firstname" value="" readonly="readonly" />
                        <label for="lastname">Last Name</label>
                         <input name="lastname" type="text" class="required" id="lastname" value="" readonly="readonly" />
                              
                         <label for="rank">Rank </label>
                       <input name="rank" type="text" class="required" id="rank" size="20" minlength="2" value="" />
                       <br/><br/><br/>
                        <label for="percent_time_app_1">Percent time appointed in JACSW: </label>
                        <input name="percent_time_app_1" type="text" class="required" id="percent_time_app_1" size="20" minlength="2" value=""/>
                        
                        <label for="org_paf">Percent time officially reallocated to grants and contracts: </label>
                        <input name="percent_time_reloc" type="text" class="required" id="percent_time_reloc" size="20" minlength="2" value=""/>
                         <br /> <br /> <br /> <br /> <br /> <br /> <br />
                        <button name="save" id="save" class="save" style="width: 60%;">Click here</button>
                        <br />
                        <br />
                    </p>
                    
                 </fieldset>
                <fieldset class="step">
                <legend> </legend>
              <br />
               <p>
                <button id="backStudents" name="backStudents" style="width: 60%;">Back to Students</button>
                 <br />
               <br />
                <label>Fall: Total Agencies <span class="small"></span>  </label>
    <input name="fall_agencies" type="text" class="required" id="fall_agencies" size="25" value=""/>
    
    <label>Total Students </label>
    <input name="fall_students" type="text" class="required" id="fall_students" size="25" value=""/>
     <br />
              <br />
               <br />
              <br />
    <label>Spring: Total Agencies </label>
    <input name="spring_agencies" type="text" class="required" id="spring_agencies" size="25" minlength="2" value=""/>
    <label>Total Students</label>
      <input name="spring_students" type="text" class="required" id="spring_students" size="25" minlength="2" value=""/>
      <br />
     
      <br />
            </p>
            
                </fieldset>
                </form>
                 <%
	%> 
            </div>
        </div>
    </div>
    <!--#include file="footer.asp"-->
</body>
</html>


