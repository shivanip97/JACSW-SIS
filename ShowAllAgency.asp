<% 
ErrMsg = Request("ErrMsg")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="Login_Check.asp"-->
<!--#include file="DBconn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>SIS | Agency</title>
<link rel="stylesheet" href="css/tabstyle.css" type="text/css" />
    <script type="text/javascript" src="jquery/jquery-1.9.0.js"></script>
    <script type="text/javascript" src="jquery/filtertable.js"></script>
	<script type="text/javascript">
	    $(document).ready(function () {



	    });
 	</script>
     <script type="text/javascript">
         function myFunctionforAgency() {
             // Declare variables 
             var input, filter, table, tr, td, i;
             input = document.getElementById("myInputAgency");
             filter = input.value.toUpperCase();
             table = document.getElementById("agencyTable");
             tr = table.getElementsByTagName("tr");

             // Loop through all table rows, and hide those who don't match the search query
             for (i = 0; i < tr.length; i++) {
                 td = tr[i].getElementsByTagName("td")[0];
                 if (td) {
                     if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {
                         tr[i].style.display = "";
                     } else {
                         tr[i].style.display = "none";
                     }
                 }
             }
         }

         function myFunctionforCity() {
             // Declare variables 
             var input, filter, table, tr, td, i;
             input = document.getElementById("myInputCity");
             filter = input.value.toUpperCase();
             table = document.getElementById("agencyTable");
             tr = table.getElementsByTagName("tr");

             // Loop through all table rows, and hide those who don't match the search query
             for (i = 0; i < tr.length; i++) {
                 td = tr[i].getElementsByTagName("td")[3];
                 if (td) {
                     if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {
                         tr[i].style.display = "";
                     } else {
                         tr[i].style.display = "none";
                     }
                 }
             }
         }
         function myFunctionforZip() {
             // Declare variables 
             var input, filter, table, tr, td, i;
             input = document.getElementById("myInputZip");
             filter = input.value.toUpperCase();
             table = document.getElementById("agencyTable");
             tr = table.getElementsByTagName("tr");

             // Loop through all table rows, and hide those who don't match the search query
             for (i = 0; i < tr.length; i++) {
                 td = tr[i].getElementsByTagName("td")[5];
                 if (td) {
                     if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {
                         tr[i].style.display = "";
                     } else {
                         tr[i].style.display = "none";
                     }
                 }
             }
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
<body bgcolor="#f2f2f2">
<!--#include file="headerAgency.asp"-->
<div align="center"><form action="" method="post"> 
     <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
     <br/> <br/>
     <div id="search">
          <p><input type="text" id="myInputAgency" onkeyup="myFunctionforAgency()" placeholder="Search for Agency"/></p>
         <br />
         <p><input type="text" id="myInputCity" onkeyup="myFunctionforCity()" placeholder="Search for City"/></p>
         <br />
         <p><input type="text" id="myInputZip" onkeyup="myFunctionforZip()" placeholder="Search for Zip"/></p>
     <br /><br />    <button type="submit" name="Button3" onclick="this.form.action='AddNewAgency.asp'; this.forms.submit();" id="Button3" value=''>Add Agency</button>
      
     
      <br /><br />
       <table id="agencyTable">
	<thead>
		<tr>
           
            <th align="center" >Agency </th>
            <th align="center">AddressL1</th>
            <th align="center">AddressL2</th>
            <th align="center">City</th>
            <th align="center">State</th>
            <th align="center">Zip</th>
           
             </tr>
        
	</thead>

    <tbody>
       
     <%
					set rs=Server.CreateObject("ADODB.recordset")
					course_query="select a.AgencyID as AgencyID, a.Agency as Agency, b.AddressL1 as AddressL1, b.AddressL2 as AddressL2, b.City as City, b.State as State, b.Zip as Zip from Agency1 a, AgencyAddress1 b where a.AgencyID = b.AgencyID  order by a.Agency"
         '
					rs.Open course_query,conn1 
                    If rs.EOF Then
                    Response.write("No Agencies Found")
                    Else
                    Do While NOT rs.Eof  
                    agencyId = rs("AgencyID")
         Agency= rs("Agency")
        AddressL1= rs("AddressL1")
        AddressL2 =rs("AddressL2")
        City= rs("City")
        State= rs("State")
        Zip= rs("Zip")
           %>
		<tr>
            
			<td align="center"><div class="edit" id="<%= agencyId%> Agency"><% Response.write (Agency) %></div></td>
            <td align="center"><div class="edit" id="<%= agencyId%> AddressL1"> <% Response.write (AddressL1) %></div></td>
            <td align="center"><div class="edit" id="<%= agencyId%> AddressL2"> <% Response.write (AddressL2) %></div></td>
            <td align="center"><div class="edit" id="<%= agencyId%> City"> <% Response.write (City) %></div></td>
            <td align="center"><div class="edit" id="<%= agencyId%> State"> <% Response.write (State) %></div></td>
            <td align="center"><div class="edit" id="<%= agencyId%> Zip"> <% Response.write (Zip) %></div></td>
           
            <td><button type="submit" name="Button1" onclick="this.form.action='ViewAgency.asp?AgencyID=<%Response.write rs("AgencyID") %>'; this.forms.submit();" id="Button1" value='<% Response.write rs("AgencyID") %>'>View/Edit Agency</button></td>
            <td><button type="submit" name="Button2" onclick="this.form.action='DeleteAgencyConfirm.asp?AgencyID=<%Response.write rs("AgencyID") %>'; this.forms.submit();" id="Button2" value='<% Response.write rs("AgencyID") %>'>Delete Agency</button></td>
         </tr>
		 <% rs.MoveNext   
        Loop End If %>
	</tbody>
    <% rs.Close
                Set rs=Nothing
                conn1.Close
                Set conn=Nothing
                %>
</table></div> 
</form> 
<!--#include file="footer.asp"-->
</div>
<!-- overlayed element -->
<div class="apple_overlay" id="overlay">
  <!-- the external content is loaded inside this tag -->
  <div class="contentWrap"></div>
</div>
<p>&nbsp;</p>

</body>
</html>