<% 
ErrMsg = Request("ErrMsg")
    AgencyID = Request("Button2")
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
<!--#include file="header.asp"-->
<!--#include file="headerAgency.asp"-->
<div align="center"><form action="" method="post"> 
     <strong><font color="#FF0000"><% Response.Write(ErrMsg) %></font></strong>
     <br/> <br/>
     <div id="search">
       <p><label for="filter">Filter</label> <input type="text" name="filter" value="" id="filter" /></p> 
     <br /><br />    <button type="submit" name="Button3" onclick="this.form.action='AddNewAgency.asp'; this.forms.submit();" id="Button3" value=''>Add Agency</button>
      </div>
      <br /><br />
       <table id="agencyTable">
	<thead>
		<tr>
           
            <th align="center">Agency</th>
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
					course_query="select a.AgencyID as AgencyID, a.Agency as Agency, b.AddressL1 as AddressL1, b.AddressL2 as AddressL2, b.City as City, b.State as State, b.Zip as Zip from Agency1 a, AgencyAddress1 b where a.AgencyID = '"& AgencyID & "' and a.AgencyID = b.AgencyID order by a.Agency"
         '
					rs.Open course_query,conn1 
                    If rs.EOF Then
                    Response.write("No Agencies Found")
                    Else
                     
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
           
            
            <td><button type="submit" name="Button2" onclick="this.form.action='DeleteAgency.asp?AgencyID=<%Response.write (AgencyID) %>'; this.forms.submit();" id="Button2" value='<% Response.write rs("AgencyID") %>'>Confirm to Delete Agency</button></td>
         </tr>
		 <% End If %>
	</tbody>
    <% rs.Close
                Set rs=Nothing
                conn1.Close
                Set conn=Nothing
                %>
</table>
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